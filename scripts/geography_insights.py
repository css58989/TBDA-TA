#!/usr/bin/env python3
"""
Geography insights (geo_questions.txt).

Reads merged-restaurant-info from Cloudflare R2 (CI) or local JSON files.
In R2 mode, also loads shoparea_and_id.xlsx from the same day prefix:
  merged-restaurant-info/year=2025/month=09/day=17/shoparea_and_id.xlsx

All answers are Excel files under --out-dir (GitHub Actions artifact in CI):
  q01_density_by_shopArea.xlsx
  q02_density_by_coordinates.xlsx
  q03_restaurant_clusters.xlsx
  q04_high_concentration_areas.xlsx
  q05_low_concentration_areas.xlsx
  q06_distribution_within_area.xlsx
  q07_distances_between_branches.xlsx
  q08_areas_concentrating_cuisines.xlsx
  q09_areas_lacking_cuisines.xlsx
  q10_many_branches_few_restaurants.xlsx
  summary.json
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from market_structure_insights import (  # noqa: E402
    COORD_OUTLIER_KM,
    DEFAULT_AREA_MAP_FILENAME,
    area_map_key_under_prefix,
    build_branch_frame,
    clean_env,
    coord_in_kuwait,
    haversine_km,
    load_area_map_local,
    load_area_map_r2,
    load_records_local,
    load_records_r2,
    make_r2_client,
    write_excel,
    write_json,
)

# Grid cell size (~1.1 km at equator; good for Kuwait relative density).
GRID_DEG = 0.01
CLUSTER_LINK_KM = 0.75
MIN_CLUSTER_SIZE = 3
TOP_N_AREAS = 25
TOP_CUISINES = 40
LACKING_MAX_SHARE = 0.02  # cuisine present in <2% of area branches counts as lacking


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def explanation_df(question_number: int, question: str, explanation: str, **extra: Any) -> pd.DataFrame:
    row = {
        "question_number": question_number,
        "question": question,
        "explanation": explanation,
        **extra,
    }
    return pd.DataFrame([row])


def annotate_geo(df: pd.DataFrame) -> pd.DataFrame:
    """Flag likely-wrong coordinates using Kuwait bbox + distance from shopArea median."""
    out = df.copy()
    area_medians: dict[Any, tuple[Optional[float], Optional[float]]] = {}
    for area_id, g in out.groupby("shopArea", dropna=False):
        coords = g.dropna(subset=["latitude", "longitude"])
        if coords.empty:
            area_medians[area_id] = (None, None)
            continue
        in_kw = coords[
            coords.apply(lambda r: coord_in_kuwait(float(r["latitude"]), float(r["longitude"])), axis=1)
        ]
        if in_kw.empty:
            area_medians[area_id] = (None, None)
        else:
            area_medians[area_id] = (
                float(in_kw["latitude"].median()),
                float(in_kw["longitude"].median()),
            )

    flags: list[str] = []
    valid: list[bool] = []
    for _, r in out.iterrows():
        lat, lon = r.get("latitude"), r.get("longitude")
        if lat is None or lon is None or (isinstance(lat, float) and math.isnan(lat)) or (
            isinstance(lon, float) and math.isnan(lon)
        ):
            flags.append("missing_coords")
            valid.append(False)
            continue
        reasons: list[str] = []
        if not coord_in_kuwait(float(lat), float(lon)):
            reasons.append("outside_kuwait_bbox")
        else:
            med_lat, med_lon = area_medians.get(r.get("shopArea"), (None, None))
            if med_lat is not None and med_lon is not None:
                dist = haversine_km(float(lat), float(lon), med_lat, med_lon)
                if dist > COORD_OUTLIER_KM:
                    reasons.append(f"outlier_vs_shopArea_median_{dist:.1f}km")
        flags.append(";".join(reasons) if reasons else "")
        valid.append(not reasons)

    out["coord_flag"] = flags
    out["coord_valid"] = valid
    med_lat_col: list[Optional[float]] = []
    med_lon_col: list[Optional[float]] = []
    for _, r in out.iterrows():
        ml, mo = area_medians.get(r.get("shopArea"), (None, None))
        med_lat_col.append(ml)
        med_lon_col.append(mo)
    out["shopArea_median_lat"] = med_lat_col
    out["shopArea_median_lon"] = med_lon_col
    return out


def coverage_km2(lats: list[float], lons: list[float]) -> Optional[float]:
    if len(lats) >= 2:
        lat_span = max(max(lats) - min(lats), 1e-5)
        lon_span = max(max(lons) - min(lons), 1e-5)
        mean_lat = sum(lats) / len(lats)
        km_lat = lat_span * 111.0
        km_lon = lon_span * 111.0 * max(0.2, abs(math.cos(math.radians(mean_lat))))
        return max(km_lat * km_lon, 0.01)
    if len(lats) == 1:
        return 0.5
    return None


def area_density_table(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        valid = g[g["coord_valid"]]
        lats = valid["latitude"].astype(float).tolist()
        lons = valid["longitude"].astype(float).tolist()
        area_km2 = coverage_km2(lats, lons)
        branches = int(g["branchId"].nunique())
        entities = int(g["entity_key"].nunique())
        dens_b = (branches / area_km2) if area_km2 else None
        dens_e = (entities / area_km2) if area_km2 else None
        rows.append(
            {
                "shopArea": area_id,
                "shopArea_name": area_name,
                "unique_branches": branches,
                "unique_restaurant_entities": entities,
                "coords_total": int(g.dropna(subset=["latitude", "longitude"]).shape[0]),
                "coords_valid": int(len(valid)),
                "coords_flagged_wrong": int((~g["coord_valid"] & g["latitude"].notna()).sum()),
                "approx_coverage_km2": round(area_km2, 4) if area_km2 is not None else None,
                "density_branches_per_km2": round(dens_b, 4) if dens_b is not None else None,
                "density_entities_per_km2": round(dens_e, 4) if dens_e is not None else None,
                "shopArea_count_rank": None,
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out = out.sort_values(
        ["density_entities_per_km2", "unique_branches"],
        ascending=[False, False],
        na_position="last",
    ).reset_index(drop=True)
    out["density_rank"] = range(1, len(out) + 1)
    out["shopArea_count_rank"] = (
        out["unique_restaurant_entities"].rank(ascending=False, method="min").astype(int)
    )
    return out


def sheet_q01(df: pd.DataFrame, density: pd.DataFrame) -> dict[str, pd.DataFrame]:
    flagged = df[~df["coord_valid"] & df["latitude"].notna()][
        [
            "shopArea",
            "shopArea_name",
            "branchId",
            "restaurantId",
            "name",
            "latitude",
            "longitude",
            "coord_flag",
        ]
    ].copy()
    by_count = density.sort_values("unique_restaurant_entities", ascending=False).copy()
    return {
        "explanation": explanation_df(
            1,
            "Restaurant density حسب المنطقة / Density by shopArea",
            (
                "Grouped by shopArea with names from shoparea_and_id.xlsx. "
                "Reports unique branches and brand-aware restaurant entities per area, "
                "plus density_entities_per_km2 from geo-validated coordinates. "
                f"Coords outside Kuwait or >{COORD_OUTLIER_KM} km from the area median are flagged wrong."
            ),
        ),
        "density_by_shopArea": density,
        "by_shopArea_count": by_count,
        "flagged_coordinates": flagged if len(flagged) else pd.DataFrame(columns=["coord_flag"]),
    }


def grid_key(lat: float, lon: float) -> tuple[float, float]:
    return (round(math.floor(lat / GRID_DEG) * GRID_DEG, 4), round(math.floor(lon / GRID_DEG) * GRID_DEG, 4))


def sheet_q02(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    valid = df[df["coord_valid"]].copy()
    rows = []
    groups: dict[tuple[float, float], list[dict[str, Any]]] = defaultdict(list)
    for _, r in valid.iterrows():
        key = grid_key(float(r["latitude"]), float(r["longitude"]))
        groups[key].append(r.to_dict())
    for (lat0, lon0), items in groups.items():
        g = pd.DataFrame(items)
        cell_km2 = (GRID_DEG * 111.0) * (GRID_DEG * 111.0 * abs(math.cos(math.radians(lat0 + GRID_DEG / 2))))
        cell_km2 = max(cell_km2, 0.01)
        branches = int(g["branchId"].nunique())
        entities = int(g["entity_key"].nunique())
        areas = (
            g.groupby(["shopArea", "shopArea_name"], dropna=False)
            .size()
            .reset_index(name="n")
            .sort_values("n", ascending=False)
        )
        top_area = areas.iloc[0]["shopArea_name"] if len(areas) else None
        rows.append(
            {
                "grid_lat_sw": lat0,
                "grid_lon_sw": lon0,
                "grid_lat_ne": round(lat0 + GRID_DEG, 4),
                "grid_lon_ne": round(lon0 + GRID_DEG, 4),
                "unique_branches": branches,
                "unique_restaurant_entities": entities,
                "density_entities_per_km2": round(entities / cell_km2, 4),
                "dominant_shopArea_name": top_area,
                "shopArea_mix_count": int(g["shopArea"].nunique()),
            }
        )
    grid = pd.DataFrame(rows)
    if not grid.empty:
        grid = grid.sort_values(
            ["density_entities_per_km2", "unique_branches"],
            ascending=[False, False],
        ).reset_index(drop=True)
        grid["rank"] = range(1, len(grid) + 1)
    return {
        "explanation": explanation_df(
            2,
            "Restaurant density حسب الإحداثيات / Density by coordinates",
            (
                f"Buckets geo-validated branches into ~{GRID_DEG}° lat/lon grid cells (~1 km). "
                "Ignores flagged-wrong coordinates. Each cell reports branch/entity counts and density. "
                "dominant_shopArea_name shows which shopArea label appears most in that cell "
                "(useful when lat/lon disagree with shopArea)."
            ),
            grid_deg=GRID_DEG,
        ),
        "density_by_grid_cell": grid if len(grid) else pd.DataFrame(),
    }


def connected_clusters(points: list[dict[str, Any]], link_km: float) -> list[list[dict[str, Any]]]:
    """Greedy connected-components clustering by haversine distance (no sklearn)."""
    n = len(points)
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri

    for i in range(n):
        for j in range(i + 1, n):
            d = haversine_km(
                points[i]["latitude"],
                points[i]["longitude"],
                points[j]["latitude"],
                points[j]["longitude"],
            )
            if d <= link_km:
                union(i, j)

    buckets: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for i, p in enumerate(points):
        buckets[find(i)].append(p)
    return list(buckets.values())


def sheet_q03(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    valid = df[df["coord_valid"]].copy()
    points = [
        {
            "branchId": r["branchId"],
            "restaurantId": r["restaurantId"],
            "entity_key": r["entity_key"],
            "name": r["name"],
            "shopArea": r["shopArea"],
            "shopArea_name": r["shopArea_name"],
            "latitude": float(r["latitude"]),
            "longitude": float(r["longitude"]),
            "cuisineString": r["cuisineString"],
        }
        for _, r in valid.iterrows()
    ]
    # Cap pairwise clustering cost: if too many points, cluster per shopArea then merge summaries.
    cluster_rows = []
    member_rows = []
    cluster_id = 0

    if len(points) <= 2500:
        groups = connected_clusters(points, CLUSTER_LINK_KM)
        source_groups = [("all", groups)]
    else:
        source_groups = []
        for area_name, g in valid.groupby("shopArea_name", dropna=False):
            area_points = [
                {
                    "branchId": r["branchId"],
                    "restaurantId": r["restaurantId"],
                    "entity_key": r["entity_key"],
                    "name": r["name"],
                    "shopArea": r["shopArea"],
                    "shopArea_name": r["shopArea_name"],
                    "latitude": float(r["latitude"]),
                    "longitude": float(r["longitude"]),
                    "cuisineString": r["cuisineString"],
                }
                for _, r in g.iterrows()
            ]
            source_groups.append((str(area_name), connected_clusters(area_points, CLUSTER_LINK_KM)))

    for scope, groups in source_groups:
        for members in groups:
            if len(members) < MIN_CLUSTER_SIZE:
                continue
            cluster_id += 1
            lats = [m["latitude"] for m in members]
            lons = [m["longitude"] for m in members]
            area_counts = Counter(m["shopArea_name"] for m in members)
            top_area, top_n = area_counts.most_common(1)[0]
            entities = {m["entity_key"] for m in members}
            cluster_rows.append(
                {
                    "cluster_id": cluster_id,
                    "scope": scope,
                    "branch_count": len(members),
                    "unique_restaurant_entities": len(entities),
                    "centroid_lat": round(sum(lats) / len(lats), 6),
                    "centroid_lon": round(sum(lons) / len(lons), 6),
                    "span_km": round(
                        max(
                            haversine_km(a["latitude"], a["longitude"], b["latitude"], b["longitude"])
                            for a in members
                            for b in members
                        ),
                        3,
                    )
                    if len(members) <= 80
                    else None,
                    "dominant_shopArea_name": top_area,
                    "dominant_shopArea_share": round(top_n / len(members), 3),
                    "shopArea_mix_count": len(area_counts),
                }
            )
            for m in members:
                member_rows.append(
                    {
                        "cluster_id": cluster_id,
                        "branchId": m["branchId"],
                        "restaurantId": m["restaurantId"],
                        "name": m["name"],
                        "shopArea": m["shopArea"],
                        "shopArea_name": m["shopArea_name"],
                        "latitude": m["latitude"],
                        "longitude": m["longitude"],
                        "cuisineString": m["cuisineString"],
                    }
                )

    clusters = pd.DataFrame(cluster_rows)
    if not clusters.empty:
        clusters = clusters.sort_values("branch_count", ascending=False).reset_index(drop=True)
    members_df = pd.DataFrame(member_rows)
    return {
        "explanation": explanation_df(
            3,
            "تجمعات المطاعم / Restaurant clusters",
            (
                f"Connected-component clusters of geo-validated branches within {CLUSTER_LINK_KM} km. "
                f"Only clusters with at least {MIN_CLUSTER_SIZE} branches are kept. "
                "dominant_shopArea_name is reported because lat/lon can disagree with shopArea."
            ),
            link_km=CLUSTER_LINK_KM,
            min_cluster_size=MIN_CLUSTER_SIZE,
        ),
        "clusters": clusters if len(clusters) else pd.DataFrame(),
        "cluster_members": members_df if len(members_df) else pd.DataFrame(),
    }


def sheet_q04_q05(density: pd.DataFrame, grid: pd.DataFrame) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame]]:
    usable = density.dropna(subset=["density_entities_per_km2"]).copy()
    ranked = usable[usable["unique_branches"] >= 3].copy()
    if ranked.empty:
        ranked = usable
    high = ranked.head(TOP_N_AREAS).copy()
    low = ranked.sort_values("density_entities_per_km2", ascending=True).head(TOP_N_AREAS).copy()

    high_grid = grid.head(TOP_N_AREAS).copy() if len(grid) else pd.DataFrame()
    low_grid = (
        grid.sort_values("density_entities_per_km2", ascending=True).head(TOP_N_AREAS).copy()
        if len(grid)
        else pd.DataFrame()
    )

    q4 = {
        "explanation": explanation_df(
            4,
            "المناطق ذات التركّز العالي / High concentration areas",
            (
                "Highest density by shopArea (geo-validated lat/lon + shopArea names) and by coordinate "
                f"grid cells. Areas need >=3 branches when available. Wrong coords excluded "
                f"(outside Kuwait or >{COORD_OUTLIER_KM} km from shopArea median)."
            ),
        ),
        "high_by_shopArea": high,
        "high_by_grid": high_grid,
    }
    q5 = {
        "explanation": explanation_df(
            5,
            "المناطق ذات التركّز المنخفض / Low concentration areas",
            (
                "Lowest density by shopArea and by coordinate grid, same geo-validation rules as Q4. "
                "Low concentration can mean sparse coverage or large approx_coverage_km2."
            ),
        ),
        "low_by_shopArea": low,
        "low_by_grid": low_grid,
    }
    return q4, q5


def sheet_q06(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    rows = []
    sample_rows = []
    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        valid = g[g["coord_valid"]]
        lats = valid["latitude"].astype(float).tolist()
        lons = valid["longitude"].astype(float).tolist()
        nn_dists = []
        for i, (la, lo) in enumerate(zip(lats, lons)):
            best = None
            for j, (lb, mb) in enumerate(zip(lats, lons)):
                if i == j:
                    continue
                d = haversine_km(la, lo, lb, mb)
                if best is None or d < best:
                    best = d
            if best is not None:
                nn_dists.append(best)

        # Simple within-area cluster count
        area_points = [
            {
                "latitude": float(r["latitude"]),
                "longitude": float(r["longitude"]),
                "branchId": r["branchId"],
            }
            for _, r in valid.iterrows()
        ]
        clusters = [c for c in connected_clusters(area_points, CLUSTER_LINK_KM) if len(c) >= 2] if area_points else []

        rows.append(
            {
                "shopArea": area_id,
                "shopArea_name": area_name,
                "unique_branches": int(g["branchId"].nunique()),
                "unique_restaurant_entities": int(g["entity_key"].nunique()),
                "coords_valid": int(len(valid)),
                "coords_flagged_wrong": int((~g["coord_valid"] & g["latitude"].notna()).sum()),
                "centroid_lat": round(sum(lats) / len(lats), 6) if lats else None,
                "centroid_lon": round(sum(lons) / len(lons), 6) if lons else None,
                "lat_span": round(max(lats) - min(lats), 6) if len(lats) >= 2 else 0.0,
                "lon_span": round(max(lons) - min(lons), 6) if len(lons) >= 2 else 0.0,
                "approx_coverage_km2": round(coverage_km2(lats, lons) or 0, 4) if lats else None,
                "avg_nearest_neighbor_km": round(sum(nn_dists) / len(nn_dists), 4) if nn_dists else None,
                "median_nearest_neighbor_km": round(float(pd.Series(nn_dists).median()), 4) if nn_dists else None,
                "within_area_cluster_count": len(clusters),
                "branches_per_restaurant_entity": round(
                    g["branchId"].nunique() / max(g["entity_key"].nunique(), 1), 4
                ),
            }
        )
        for _, r in valid.head(30).iterrows():
            sample_rows.append(
                {
                    "shopArea": area_id,
                    "shopArea_name": area_name,
                    "branchId": r["branchId"],
                    "name": r["name"],
                    "latitude": r["latitude"],
                    "longitude": r["longitude"],
                    "cuisineString": r["cuisineString"],
                }
            )

    dist = pd.DataFrame(rows)
    if not dist.empty:
        dist = dist.sort_values("unique_branches", ascending=False).reset_index(drop=True)
    return {
        "explanation": explanation_df(
            6,
            "توزيع المطاعم داخل المنطقة / Distribution within area",
            (
                "Per shopArea: centroid, lat/lon span, coverage km², nearest-neighbor spacing, "
                f"and within-area cluster count (link {CLUSTER_LINK_KM} km). "
                "Uses geo-validated coordinates only for spatial metrics."
            ),
        ),
        "within_area_summary": dist if len(dist) else pd.DataFrame(),
        "sample_points_per_area": pd.DataFrame(sample_rows),
    }


def sheet_q07(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Distances between branches of the same restaurant entity (multi-branch only)."""
    valid = df[df["coord_valid"]].copy()
    entity_stats = []
    pair_samples = []

    for entity_key, g in valid.groupby("entity_key"):
        if g["branchId"].nunique() < 2:
            continue
        pts = g.drop_duplicates("branchId")
        coords = [
            (
                r["branchId"],
                r["name"],
                r["branchName"],
                r["shopArea_name"],
                float(r["latitude"]),
                float(r["longitude"]),
            )
            for _, r in pts.iterrows()
        ]
        dists = []
        for i in range(len(coords)):
            for j in range(i + 1, len(coords)):
                d = haversine_km(coords[i][4], coords[i][5], coords[j][4], coords[j][5])
                dists.append(d)
                if len(pair_samples) < 5000:
                    pair_samples.append(
                        {
                            "entity_key": entity_key,
                            "display_name": coords[i][1],
                            "branchId_a": coords[i][0],
                            "branchId_b": coords[j][0],
                            "shopArea_a": coords[i][3],
                            "shopArea_b": coords[j][3],
                            "distance_km": round(d, 4),
                        }
                    )
        if not dists:
            continue
        nn = []
        for i in range(len(coords)):
            best = min(
                haversine_km(coords[i][4], coords[i][5], coords[j][4], coords[j][5])
                for j in range(len(coords))
                if j != i
            )
            nn.append(best)
        entity_stats.append(
            {
                "entity_key": entity_key,
                "display_name": str(pts["name"].mode().iloc[0]) if len(pts["name"].dropna()) else entity_key,
                "special_franchise": pts["special_franchise"].dropna().iloc[0]
                if pts["special_franchise"].notna().any()
                else None,
                "branch_count_with_valid_coords": len(coords),
                "avg_pairwise_distance_km": round(sum(dists) / len(dists), 4),
                "median_pairwise_distance_km": round(float(pd.Series(dists).median()), 4),
                "min_pairwise_distance_km": round(min(dists), 4),
                "max_pairwise_distance_km": round(max(dists), 4),
                "avg_nearest_sibling_km": round(sum(nn) / len(nn), 4),
                "shopArea_count": int(pts["shopArea"].nunique()),
            }
        )

    entity_df = pd.DataFrame(entity_stats)
    if not entity_df.empty:
        entity_df = entity_df.sort_values(
            ["branch_count_with_valid_coords", "avg_pairwise_distance_km"],
            ascending=[False, False],
        ).reset_index(drop=True)

    pairs = pd.DataFrame(pair_samples)
    if not pairs.empty:
        pairs = pairs.sort_values("distance_km", ascending=True).reset_index(drop=True)

    # Cross-brand nearest neighbor (market spacing)
    market_nn = []
    pts_all = valid.drop_duplicates("branchId")
    # Sample if huge
    if len(pts_all) > 3000:
        pts_all = pts_all.sample(n=3000, random_state=42)
    coords_all = [
        (r["branchId"], r["name"], float(r["latitude"]), float(r["longitude"]), r["shopArea_name"])
        for _, r in pts_all.iterrows()
    ]
    for i, (bid, name, la, lo, area) in enumerate(coords_all):
        best = None
        best_other = None
        for j, (bid2, name2, lb, mb, area2) in enumerate(coords_all):
            if i == j:
                continue
            d = haversine_km(la, lo, lb, mb)
            if best is None or d < best:
                best = d
                best_other = (bid2, name2, area2)
        if best is not None and best_other is not None:
            market_nn.append(
                {
                    "branchId": bid,
                    "name": name,
                    "shopArea_name": area,
                    "nearest_other_branchId": best_other[0],
                    "nearest_other_name": best_other[1],
                    "nearest_other_shopArea": best_other[2],
                    "nearest_distance_km": round(best, 4),
                }
            )
    nn_df = pd.DataFrame(market_nn)
    if not nn_df.empty:
        nn_df = nn_df.sort_values("nearest_distance_km").reset_index(drop=True)

    return {
        "explanation": explanation_df(
            7,
            "المسافات بين الفروع / Distances between branches",
            (
                "For multi-branch restaurant entities (brand-aware: McDonald's/McCafe/Starbucks collapsed), "
                "pairwise and nearest-sibling distances using geo-validated coordinates. "
                "Also nearest-neighbor spacing across all branches (market packing). "
                "Pairwise sheet may be capped for size."
            ),
        ),
        "per_restaurant_entity": entity_df if len(entity_df) else pd.DataFrame(),
        "pairwise_samples": pairs if len(pairs) else pd.DataFrame(),
        "market_nearest_neighbor": nn_df if len(nn_df) else pd.DataFrame(),
    }


def cuisine_tags(raw: object) -> list[str]:
    text = str(raw or "").strip()
    parts = [p.strip() for p in text.split(",") if p.strip()]
    return parts or ["(blank)"]


def sheet_q08_q09(df: pd.DataFrame) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame]]:
    # Global cuisine popularity (by unique branches)
    global_cuisine_branches: dict[str, set[Any]] = defaultdict(set)
    for _, r in df.iterrows():
        for tag in cuisine_tags(r["cuisineString"]):
            global_cuisine_branches[tag].add(r["branchId"])
    top_cuisines = [
        c for c, _ in sorted(global_cuisine_branches.items(), key=lambda kv: -len(kv[1]))[:TOP_CUISINES]
    ]

    concentrate_rows = []
    lacking_rows = []
    area_cuisine_matrix = []

    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        branch_n = max(int(g["branchId"].nunique()), 1)
        area_cuisine: dict[str, set[Any]] = defaultdict(set)
        for _, r in g.iterrows():
            for tag in cuisine_tags(r["cuisineString"]):
                area_cuisine[tag].add(r["branchId"])

        for cuisine in top_cuisines:
            local_n = len(area_cuisine.get(cuisine, set()))
            share = local_n / branch_n
            global_n = len(global_cuisine_branches[cuisine])
            lift = (share / (global_n / max(int(df["branchId"].nunique()), 1))) if global_n else None
            area_cuisine_matrix.append(
                {
                    "shopArea": area_id,
                    "shopArea_name": area_name,
                    "cuisine": cuisine,
                    "branches_with_cuisine": local_n,
                    "area_branches": branch_n,
                    "share_of_area": round(share, 4),
                    "lift_vs_global": round(lift, 4) if lift is not None else None,
                }
            )
            if local_n >= 3 and lift is not None and lift >= 1.5:
                concentrate_rows.append(
                    {
                        "shopArea": area_id,
                        "shopArea_name": area_name,
                        "cuisine": cuisine,
                        "branches_with_cuisine": local_n,
                        "share_of_area": round(share, 4),
                        "lift_vs_global": round(lift, 4),
                    }
                )
            if local_n == 0 or share <= LACKING_MAX_SHARE:
                # Only flag as lacking if cuisine is common globally
                if global_n >= 20:
                    lacking_rows.append(
                        {
                            "shopArea": area_id,
                            "shopArea_name": area_name,
                            "cuisine": cuisine,
                            "branches_with_cuisine": local_n,
                            "share_of_area": round(share, 4),
                            "global_branches_with_cuisine": global_n,
                            "status": "absent" if local_n == 0 else "rare",
                        }
                    )

    conc = pd.DataFrame(concentrate_rows)
    if not conc.empty:
        conc = conc.sort_values(["lift_vs_global", "branches_with_cuisine"], ascending=[False, False]).reset_index(
            drop=True
        )
    lack = pd.DataFrame(lacking_rows)
    if not lack.empty:
        lack = lack.sort_values(
            ["global_branches_with_cuisine", "share_of_area"],
            ascending=[False, True],
        ).reset_index(drop=True)
    matrix = pd.DataFrame(area_cuisine_matrix)

    q8 = {
        "explanation": explanation_df(
            8,
            "المناطق التي تجمع أنواعًا معينة من المطاعم / Areas concentrating cuisine types",
            (
                f"For the top {TOP_CUISINES} cuisine tags globally, finds shopAreas where the cuisine "
                "share is at least 1.5× the global share (lift) and appears on >=3 branches. "
                "Uses cuisineString tags (comma-split)."
            ),
        ),
        "concentrated_cuisines": conc if len(conc) else pd.DataFrame(),
        "area_cuisine_matrix_top": matrix.sort_values("lift_vs_global", ascending=False).head(5000)
        if len(matrix)
        else pd.DataFrame(),
    }
    q9 = {
        "explanation": explanation_df(
            9,
            "المناطق التي تفتقر إلى أنواع معينة من المطاعم / Areas lacking cuisine types",
            (
                f"shopAreas where a globally common cuisine (>=20 branches) is absent or "
                f"<= {LACKING_MAX_SHARE:.0%} of local branches. Helps spot white-space / gaps."
            ),
        ),
        "lacking_cuisines": lack if len(lack) else pd.DataFrame(),
    }
    return q8, q9


def sheet_q10(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    rows = []
    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        branches = int(g["branchId"].nunique())
        entities = int(g["entity_key"].nunique())
        raw_rids = int(g["restaurantId"].nunique())
        if entities == 0:
            continue
        ratio = branches / entities
        rows.append(
            {
                "shopArea": area_id,
                "shopArea_name": area_name,
                "unique_branches": branches,
                "unique_restaurant_entities": entities,
                "unique_restaurantIds_raw": raw_rids,
                "branches_per_restaurant_entity": round(ratio, 4),
                "extra_branches_beyond_one_each": branches - entities,
                "pct_branches_from_multi_branch_entities": round(
                    100.0
                    * (
                        g.groupby("entity_key")["branchId"].nunique().pipe(lambda s: s[s > 1].sum())
                    )
                    / branches,
                    2,
                )
                if branches
                else None,
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        # Prefer areas that both have many branches and high branch/entity ratio
        out["many_branches_few_restaurants_score"] = out["unique_branches"] * out[
            "branches_per_restaurant_entity"
        ]
        out = out.sort_values(
            ["many_branches_few_restaurants_score", "branches_per_restaurant_entity"],
            ascending=[False, False],
        ).reset_index(drop=True)
        out["rank"] = range(1, len(out) + 1)
    return {
        "explanation": explanation_df(
            10,
            "المناطق التي فيها عدد كبير من الفروع لكن عدد قليل من المطاعم الفعلية",
            (
                "Per shopArea: unique_branches vs brand-aware unique_restaurant_entities. "
                "High branches_per_restaurant_entity means chains dominate (many branches, fewer brands). "
                "McDonald's/McCafe/Starbucks are collapsed to one entity each."
            ),
        ),
        "by_shopArea": out if len(out) else pd.DataFrame(),
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Geography insights (geo_questions.txt)")
    p.add_argument("--mode", choices=["r2", "local"], default=os.environ.get("INSIGHTS_MODE", "r2"))
    p.add_argument("--local-dir", type=Path, default=None)
    p.add_argument(
        "--prefix",
        default=os.environ.get(
            "R2_PREFIX",
            "merged-restaurant-info/year=2025/month=09/day=17/",
        ),
    )
    p.add_argument(
        "--key-contains",
        default=os.environ.get("R2_KEY_CONTAINS", "MergedRestaurantsInfo.json"),
    )
    p.add_argument("--max-files", type=int, default=None)
    p.add_argument("--area-map", type=Path, default=None)
    p.add_argument("--area-map-key", default=os.environ.get("AREA_MAP_R2_KEY", ""))
    p.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "outputs" / "geography",
        help="Local staging dir (gitignored). In CI uploaded as a GitHub Actions artifact.",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()
    area_map_source = ""

    if args.mode == "local":
        local_dir = args.local_dir
        if local_dir is None:
            candidate = ROOT.parent / "Restaurant_info"
            local_dir = candidate if candidate.exists() else ROOT / "sample_data"
        if not local_dir.exists():
            raise SystemExit(f"Local directory not found: {local_dir}")

        area_path = args.area_map
        if area_path is None:
            for candidate in (
                local_dir / DEFAULT_AREA_MAP_FILENAME,
                ROOT / "data" / DEFAULT_AREA_MAP_FILENAME,
                ROOT.parent / DEFAULT_AREA_MAP_FILENAME,
            ):
                if candidate.exists():
                    area_path = candidate
                    break
        if area_path is None:
            print(
                f"[warn] {DEFAULT_AREA_MAP_FILENAME} not found locally; "
                "using unknown_area_<id> labels. Pass --area-map when available."
            )
            area_map = {}
            area_map_source = "(missing local area map)"
        else:
            area_map = load_area_map_local(area_path)
            area_map_source = str(area_path)
        records, sources = load_records_local(local_dir, args.key_contains, args.max_files)
        mode = "local"
    else:
        bucket = clean_env("CF_R2_BUCKET_NAME")
        client = make_r2_client(bucket=bucket)
        area_key = (args.area_map_key or "").strip() or area_map_key_under_prefix(args.prefix)
        area_map = load_area_map_r2(client, bucket, area_key)
        area_map_source = f"s3://{bucket}/{area_key}"
        records, sources = load_records_r2(
            client, bucket, args.prefix, args.key_contains, args.max_files
        )
        mode = "r2"

    if not records:
        raise SystemExit("No restaurant records loaded")

    df = build_branch_frame(records, area_map)
    df = annotate_geo(df)
    density = area_density_table(df)
    out = args.out_dir
    out.mkdir(parents=True, exist_ok=True)

    write_excel(out / "q01_density_by_shopArea.xlsx", sheet_q01(df, density))
    q02 = sheet_q02(df)
    write_excel(out / "q02_density_by_coordinates.xlsx", q02)
    write_excel(out / "q03_restaurant_clusters.xlsx", sheet_q03(df))
    q04, q05 = sheet_q04_q05(density, q02["density_by_grid_cell"])
    write_excel(out / "q04_high_concentration_areas.xlsx", q04)
    write_excel(out / "q05_low_concentration_areas.xlsx", q05)
    write_excel(out / "q06_distribution_within_area.xlsx", sheet_q06(df))
    write_excel(out / "q07_distances_between_branches.xlsx", sheet_q07(df))
    q08, q09 = sheet_q08_q09(df)
    write_excel(out / "q08_areas_concentrating_cuisines.xlsx", q08)
    write_excel(out / "q09_areas_lacking_cuisines.xlsx", q09)
    write_excel(out / "q10_many_branches_few_restaurants.xlsx", sheet_q10(df))

    summary = {
        "generated_at": utc_now(),
        "mode": mode,
        "source_files": sources,
        "raw_rows_loaded": len(records),
        "unique_branches_after_dedupe": int(df["branchId"].nunique()),
        "coords_valid": int(df["coord_valid"].sum()),
        "coords_flagged_wrong": int((~df["coord_valid"] & df["latitude"].notna()).sum()),
        "area_map_source": area_map_source,
        "excel": [
            "q01_density_by_shopArea.xlsx",
            "q02_density_by_coordinates.xlsx",
            "q03_restaurant_clusters.xlsx",
            "q04_high_concentration_areas.xlsx",
            "q05_low_concentration_areas.xlsx",
            "q06_distribution_within_area.xlsx",
            "q07_distances_between_branches.xlsx",
            "q08_areas_concentrating_cuisines.xlsx",
            "q09_areas_lacking_cuisines.xlsx",
            "q10_many_branches_few_restaurants.xlsx",
        ],
        "reference_inputs": [
            "data/geo_questions.txt",
            area_map_source,
        ],
    }
    write_json(out / "summary.json", summary)
    print("")
    print(
        f"Done: branches={summary['unique_branches_after_dedupe']}, "
        f"valid_coords={summary['coords_valid']}, "
        f"flagged={summary['coords_flagged_wrong']}, "
        f"files={len(sources)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
