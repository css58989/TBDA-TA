#!/usr/bin/env python3
"""
Branch-level geography insights (city / area distribution, density, highest / lowest areas).

Why this exists
---------------
`unique_restaurants_by_id` keeps ONE row per restaurant id, so shopCity / shopArea of the
first-seen branch are copied to every branch of that restaurant (only ~38% were right).
The raw MergedRestaurantsInfo.json records are already per branch: every record has its own
branchId, shopArea, shopCity, latitude, longitude. The same branch shows up in several area
files (the file = the delivery area that was scraped), so we de-duplicate by branchId and keep
the branch's OWN shopArea / shopCity.

Coordinates are then used to validate shopArea (coordinates are sometimes wrong too):
  1. Coordinates outside Kuwait are flagged and ignored (swapped lat/lon are repaired).
  2. Reference points = branches whose own shopArea equals the area written in branchName.
  3. kNN on reference points gives the area implied by the coordinates.
  4. Decision per branch (area_basis column):
       shopArea confirmed by coords / by branch name
       corrected  (coords AND branch name agree on another area)
       kept_suspect (coords disagree with shopArea but nothing else confirms) -> shopArea kept
       kept_no_evidence

Outputs (each answer is its own Excel file) in --out-dir:
  q00_branches.xlsx                      one row per branch (requested columns + validation)
  q07_branches_by_city.xlsx              branches / restaurants per shopCity
  q08_branch_density_by_area.xlsx        branches, restaurants, km2, branches per km2 per shopArea
  q09_highest_lowest_density_areas.xlsx  top / bottom areas by count and by density
  q_area_validation.xlsx                 branches whose shopArea was not confirmed
  summary.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.restaurant_io import (  # noqa: E402
    clean_env,
    iter_local_files,
    list_r2_keys,
    load_json_records_from_bytes,
    make_r2_client,
)

AREA_MAP_FILENAME = "shoparea_and_id.xlsx"

# City names are not in the data; inferred from the areas inside each shopCity id.
CITY_NAMES = {
    1: "Al Asimah (Capital)",
    2: "Hawalli",
    3: "Farwaniya",
    4: "Ahmadi",
    5: "Jahra",
    6: "Mubarak Al-Kabeer",
}

KEEP = [
    "branchId", "branchName", "restaurantId", "id", "name", "latitude", "longitude",
    "shopCity", "shopArea", "cuisineString", "isGrocery", "verticalType", "shopType", "isDarkstore",
]

KUWAIT = dict(lat=(28.5, 30.3), lon=(46.4, 48.8))


def norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def load_area_map(df: pd.DataFrame) -> dict[int, str]:
    return {int(r.area_id): str(r.area_name) for r in df.itertuples() if pd.notna(r.area_id)}


def slim(rec: dict, source: str) -> dict:
    out = {k: rec.get(k) for k in KEEP}
    if out["restaurantId"] in (None, ""):
        out["restaurantId"] = out["id"]
    out["source_file"] = source
    return out


def collect(args) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows: list[dict] = []
    if args.mode == "local":
        area_df = pd.read_excel(Path(args.area_map))
        files = list(iter_local_files(Path(args.local_dir), args.key_contains))
        if args.max_files:
            files = files[: args.max_files]
        for i, (key, path) in enumerate(files, 1):
            recs = load_json_records_from_bytes(path.read_bytes())
            rows += [slim(r, Path(key).name) for r in recs]
            print(f"[{i}/{len(files)}] {Path(key).name}: {len(recs)} records")
    else:
        client = make_r2_client()
        bucket = clean_env("CF_R2_BUCKET_NAME")
        prefix = args.prefix if args.prefix.endswith("/") or not args.prefix else args.prefix + "/"
        obj = client.get_object(Bucket=bucket, Key=f"{prefix}{AREA_MAP_FILENAME}")
        area_df = pd.read_excel(BytesIO(obj["Body"].read()))
        keys = list_r2_keys(client, bucket, args.prefix, args.key_contains)
        if args.max_files:
            keys = keys[: args.max_files]
        for i, key in enumerate(keys, 1):
            body = client.get_object(Bucket=bucket, Key=key)["Body"].read()
            recs = load_json_records_from_bytes(body)
            rows += [slim(r, key.split("/")[-1]) for r in recs]
            print(f"[{i}/{len(keys)}] {key}: {len(recs)} records")
    if not rows:
        raise SystemExit("No records loaded.")
    return pd.DataFrame(rows), area_df


def dedupe_branches(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw[raw.branchId.notna()].copy()
    raw["latitude"] = pd.to_numeric(raw.latitude, errors="coerce")
    raw["longitude"] = pd.to_numeric(raw.longitude, errors="coerce")
    n_files = raw.groupby("branchId").source_file.nunique().rename("seen_in_area_files")
    # same branch in many files: keep the most frequent shopArea/shopCity, first record for the rest
    mode_area = raw.groupby("branchId").shopArea.agg(lambda s: s.mode().iloc[0])
    mode_city = raw.groupby("branchId").shopCity.agg(lambda s: s.mode().iloc[0])
    conflict = raw.groupby("branchId").shopArea.nunique().rename("n_distinct_shopArea")
    b = raw.drop_duplicates("branchId").set_index("branchId")
    b["shopArea"] = mode_area
    b["shopCity"] = mode_city
    b = b.join(n_files).join(conflict).reset_index()
    return b


def validate_and_correct(b: pd.DataFrame, id2name: dict[int, str]) -> pd.DataFrame:
    from sklearn.neighbors import KNeighborsClassifier, NearestNeighbors

    key2id: dict[str, int] = {}
    for aid, an in id2name.items():
        key2id.setdefault(norm(re.sub(r"^\d+-", "", an)), aid)

    lat_ok = b.latitude.between(*KUWAIT["lat"])
    lon_ok = b.longitude.between(*KUWAIT["lon"])
    swapped = (~(lat_ok & lon_ok)) & b.longitude.between(*KUWAIT["lat"]) & b.latitude.between(*KUWAIT["lon"])
    b.loc[swapped, ["latitude", "longitude"]] = b.loc[swapped, ["longitude", "latitude"]].values
    b["coords_swapped_fixed"] = swapped
    b["coords_valid"] = b.latitude.between(*KUWAIT["lat"]) & b.longitude.between(*KUWAIT["lon"])

    def suffix(n):
        s = str(n).split(",")[-1].strip() if "," in str(n) else ""
        s = re.sub(r"\s*(TGO|TMP)\s*$", "", s, flags=re.I)
        return re.sub(r"^(tgo|tmp)\s+", "", s, flags=re.I)

    b["name_area_id"] = b.branchName.map(suffix).map(lambda s: key2id.get(norm(s)))
    b["shopArea"] = pd.to_numeric(b.shopArea, errors="coerce")

    ref = b[b.coords_valid & (b.name_area_id == b.shopArea)]
    vc = ref.shopArea.value_counts()
    ref = ref[ref.shopArea.isin(vc[vc >= 3].index)]
    cols = ["latitude", "longitude"]
    b["coord_area_id"] = np.nan
    if len(ref) >= 10:
        knn = KNeighborsClassifier(n_neighbors=min(7, len(ref)), weights="distance").fit(ref[cols], ref.shopArea)
        v = b.coords_valid
        b.loc[v, "coord_area_id"] = knn.predict(b.loc[v, cols])
        dist, _ = NearestNeighbors(n_neighbors=1).fit(ref[cols]).kneighbors(b.loc[v, cols])
        far = pd.Series(dist[:, 0] * 111 > 3, index=b.index[v])
        b.loc[far[far].index, "coord_area_id"] = np.nan

    def decide(r):
        s, n, c = r.shopArea, r.name_area_id, r.coord_area_id
        n = None if pd.isna(n) else n
        c = None if pd.isna(c) else c
        if c is not None and n is not None and c == n:
            return (s, "shopArea confirmed by coords + name") if c == s else (c, "corrected: coords + name agree on other area")
        if c is not None and c == s:
            return s, "shopArea confirmed by coords"
        if n is not None and n == s:
            return s, "shopArea confirmed by branch name"
        if c is not None:
            return s, "kept_suspect: coords imply another area"
        return s, "kept_no_evidence"

    res = b.apply(decide, axis=1, result_type="expand")
    b["final_area_id"] = res[0].astype("Int64")
    b["area_basis"] = res[1]
    b["area_changed"] = b.final_area_id != b.shopArea
    area_city = b.groupby("shopArea").shopCity.agg(lambda s: s.mode().iloc[0])
    b["final_city_id"] = b.final_area_id.map(area_city).fillna(b.shopCity)
    return b


def name_of(aid, id2name):
    return id2name.get(int(aid), f"area-{int(aid)}") if pd.notna(aid) else None


def hull_km2(g: pd.DataFrame) -> float:
    from scipy.spatial import ConvexHull

    p = g.loc[g.coords_valid, ["latitude", "longitude"]].values
    if len(p) < 5:
        return np.nan
    d = np.linalg.norm(p - np.median(p, axis=0), axis=1)
    p = p[d <= np.quantile(d, 0.9)]
    x = (p[:, 1] - p[:, 1].mean()) * 111 * np.cos(np.radians(29.3))
    y = (p[:, 0] - p[:, 0].mean()) * 111
    try:
        return float(ConvexHull(np.c_[x, y]).volume)
    except Exception:
        return np.nan


def build_outputs(b: pd.DataFrame, id2name: dict[int, str], out: Path, raw_rows: int, sources: int) -> dict:
    out.mkdir(parents=True, exist_ok=True)
    cname = lambda c: CITY_NAMES.get(int(c), f"city-{int(c)}") if pd.notna(c) else None
    b["shopCity_name"] = b.shopCity.map(cname)
    b["shopArea_name"] = b.shopArea.map(lambda a: name_of(a, id2name))
    b["final_area_name"] = b.final_area_id.map(lambda a: name_of(a, id2name))
    b["final_city_name"] = b.final_city_id.map(cname)
    rname = b.name.fillna("").str.lower().str.strip()
    b["_rname"] = rname.where(rname != "", b.restaurantId.astype(str))

    branches = b[[
        "branchId", "branchName", "restaurantId", "name", "longitude", "latitude",
        "shopCity", "shopCity_name", "shopArea", "shopArea_name",
        "cuisineString", "isGrocery", "verticalType", "shopType", "isDarkstore",
        "final_city_id", "final_city_name", "final_area_id", "final_area_name",
        "area_basis", "coords_valid", "seen_in_area_files",
    ]].rename(columns={
        "branchId": "branch_id", "branchName": "branch_name", "restaurantId": "restaurant_id",
        "name": "restaurant_name", "shopCity": "shop_city_id", "shopCity_name": "shop_city_name",
        "shopArea": "shop_area_id", "shopArea_name": "shop_area_name",
    })
    branches.to_excel(out / "q00_branches.xlsx", index=False)

    # Q7
    def city_tab(idc, namec):
        g = b.groupby([idc, namec]).agg(
            branches=("branchId", "nunique"),
            restaurant_ids=("restaurantId", "nunique"),
            restaurants_by_name=("_rname", "nunique"),
        ).reset_index()
        g["pct_of_branches"] = (g.branches / g.branches.sum() * 100).round(2)
        return g.sort_values("branches", ascending=False)
    q7 = city_tab("shopCity", "shopCity_name").rename(columns={"shopCity": "city_id", "shopCity_name": "city_name"})
    q7c = city_tab("final_city_id", "final_city_name").rename(columns={"final_city_id": "city_id", "final_city_name": "city_name"})
    with pd.ExcelWriter(out / "q07_branches_by_city.xlsx") as w:
        q7.to_excel(w, sheet_name="by_shopCity", index=False)
        q7c.to_excel(w, sheet_name="after_coord_validation", index=False)

    # Q8
    rows = []
    for aid, g in b.groupby("final_area_id"):
        raw = b[b.shopArea == aid]
        rows.append(dict(
            area_id=int(aid), area_name=name_of(aid, id2name),
            city_name=g.final_city_name.mode().iloc[0] if g.final_city_name.notna().any() else None,
            branches=len(g), restaurants_by_name=g._rname.nunique(), restaurant_ids=g.restaurantId.nunique(),
            branches_by_raw_shopArea=len(raw),
            branches_with_valid_coords=int(g.coords_valid.sum()),
            pct_confirmed_by_coords_or_name=round(g.area_basis.str.contains("confirmed").mean() * 100, 1),
            approx_area_km2=hull_km2(g),
        ))
    q8 = pd.DataFrame(rows)
    q8["approx_area_km2"] = q8.approx_area_km2.clip(lower=0.25).round(2)
    q8["branches_per_km2"] = (q8.branches / q8.approx_area_km2).round(1)
    q8["note"] = np.where(q8.approx_area_km2.isna(), "fewer than 5 valid coordinates - density not computed", "")
    q8 = q8.sort_values("branches", ascending=False)
    q8.insert(0, "rank_by_branches", range(1, len(q8) + 1))
    q8.to_excel(out / "q08_branch_density_by_area.xlsx", index=False)

    # Q9
    ok = q8[q8.branches_per_km2.notna()]
    parts = [
        q8.head(15).assign(list="Highest by branch count"),
        q8.tail(15).iloc[::-1].assign(list="Lowest by branch count"),
        ok.sort_values("branches_per_km2", ascending=False).head(15).assign(list="Highest by density (branches/km2)"),
        ok.sort_values("branches_per_km2").head(15).assign(list="Lowest by density (branches/km2)"),
    ]
    q9 = pd.concat(parts)
    q9 = q9[["list"] + [c for c in q9.columns if c != "list"]]
    q9.to_excel(out / "q09_highest_lowest_density_areas.xlsx", index=False)

    # validation
    flagged = b[~b.area_basis.str.contains("confirmed")][[
        "branchId", "branchName", "name", "latitude", "longitude", "shopArea", "shopArea_name",
        "coord_area_id", "name_area_id", "final_area_id", "final_area_name", "area_basis", "coords_valid",
    ]]
    flagged.to_excel(out / "q_area_validation.xlsx", index=False)

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "raw_records": raw_rows, "source_files": sources,
        "unique_branches": int(len(b)),
        "unique_restaurant_ids": int(b.restaurantId.nunique()),
        "unique_restaurants_by_name": int(b._rname.nunique()),
        "branches_invalid_coordinates": int((~b.coords_valid).sum()),
        "branches_swapped_coords_fixed": int(b.coords_swapped_fixed.sum()),
        "branches_with_conflicting_shopArea_across_files": int((b.n_distinct_shopArea > 1).sum()),
        "area_basis_counts": b.area_basis.value_counts().to_dict(),
        "areas_corrected": int(b.area_changed.sum()),
        "city_names_note": "inferred: 1 Capital, 2 Hawalli, 3 Farwaniya, 4 Ahmadi, 5 Jahra, 6 Mubarak Al-Kabeer",
    }
    (out / "summary.json").write_text(json.dumps(summary, indent=2, default=str), encoding="utf-8")
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["local", "r2"], default="local")
    ap.add_argument("--local-dir", default=str(ROOT.parent / "Restaurant_info"))
    ap.add_argument("--area-map", default=str(ROOT / "data" / AREA_MAP_FILENAME))
    ap.add_argument("--prefix", default="merged-restaurant-info/year=2025/month=09/day=17/")
    ap.add_argument("--key-contains", default="MergedRestaurantsInfo.json")
    ap.add_argument("--max-files", type=int, default=0)
    ap.add_argument("--out-dir", default=str(ROOT / "outputs" / "branch_geography"))
    args = ap.parse_args()

    raw, area_df = collect(args)
    id2name = load_area_map(area_df)
    b = dedupe_branches(raw)
    b = validate_and_correct(b, id2name)
    s = build_outputs(b, id2name, Path(args.out_dir), len(raw), raw.source_file.nunique())
    print(json.dumps(s, indent=2, default=str))


if __name__ == "__main__":
    main()
