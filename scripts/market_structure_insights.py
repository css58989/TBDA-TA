#!/usr/bin/env python3
"""
Market / Restaurant Structure insights (insightspart1.txt questions 1-11).

Reads merged-restaurant-info from Cloudflare R2 (CI) or local JSON files.
In R2 mode, also loads shoparea_and_id.xlsx from the same day prefix:
  merged-restaurant-info/year=2025/month=09/day=17/shoparea_and_id.xlsx

Writes per-question files under --out-dir (uploaded as a GitHub Actions artifact in CI):
  - JSON with explanation for Q1, Q2, Q5
  - Excel for Q1-Q4, Q6-Q11 (Q1/Q2 include id+name listing sheets)

Restaurant entity keys (used for actual_unique across all questions):
  - unique_by_restaurant_id = distinct restaurantId
  - actual_unique_restaurant_entities = distinct normalized restaurant name
    (area suffixes stripped, e.g. "Asha's, Jabriya" -> "asha's"), with
    McDonald's / McCafe / Starbucks forced to one brand key each when the
    name does not use a clean comma pattern.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

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

# Brands where each physical branch gets a distinct restaurantId.
SPECIAL_FRANCHISE_PATTERNS: list[tuple[str, str]] = [
    (r"mcdonald", "McDonald's"),
    (r"mccafe", "McCafe"),
    (r"starbucks", "Starbucks"),
]

VERTICAL_TYPE_NAMES: dict[int, str] = {
    0: "restaurants/food",
    1: "grocery/supermarket",
    2: "pharmacy/health",
    3: "flowers",
    4: "electronics",
    5: "pet shops",
    6: "cosmetics/beauty",
    9: "specialty stores",
}

# Inferred from dominant shopArea clusters in Kuwait data (numeric shopCity is source of truth).
SHOPCITY_LABELS: dict[int, str] = {
    1: "Capital / Kuwait City",
    2: "Hawalli",
    3: "Farwaniya",
    4: "Ahmadi",
    5: "Jahra",
    6: "Mubarak Al-Kabeer",
    8: "Other / Unmapped",
}

# Approximate Kuwait bounding box for coordinate sanity checks.
KUWAIT_LAT = (28.5, 30.2)
KUWAIT_LON = (46.5, 48.8)
COORD_OUTLIER_KM = 25.0


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


_APOSTROPHE_RE = re.compile(r"[\u2018\u2019\u02bc`´]")
_PAREN_SUFFIX_RE = re.compile(
    r"\s*\((?:DH\s*Kitchen|TGO|Talabat\s*GO)[^)]*\)\s*",
    flags=re.IGNORECASE,
)
# Text after " - " that looks like a branch/location (not part of the brand name).
_BRANCH_LOCATION_HINTS: tuple[str, ...] = (
    "mall",
    "tower",
    "kiosk",
    "club",
    "co-op",
    "co op",
    "kitchen",
    "restaurants",
    "land ",
    "land,",
    "gate",
    "lake",
    "fanar",
    "avenues",
    "360",
    "jahra",
    "salmiya",
    "jabriya",
    "hawall",
    "farwaniya",
    "mahboula",
    "egal",
    "zahra",
    "rai",
    "shuwaikh",
    "ardhiya",
    "subhan",
    "mirqab",
    "salwa",
    "fnaitees",
    "egaliya",
    "egal",
    "abdullah",
    "salem",
    "mourouj",
    "nasser",
    "sidra",
    "bneid",
    "reggai",
    "sabah",
)


def _clean_text(value: object) -> Optional[str]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    text = _APOSTROPHE_RE.sub("'", str(value)).strip()
    return text or None


def detect_special_franchise(name: object, branch_name: object = None) -> Optional[str]:
    for value in (name, branch_name):
        text = _clean_text(value)
        if text is None:
            continue
        lowered = text.lower()
        for pattern, label in SPECIAL_FRANCHISE_PATTERNS:
            if re.search(pattern, lowered):
                return label
    return None


def _looks_like_branch_location_suffix(suffix: str) -> bool:
    s = suffix.lower().strip(" ,.-_")
    if not s:
        return False
    if s in {"tgo", "dh kitchen", "talabat go"}:
        return True
    return any(hint in s for hint in _BRANCH_LOCATION_HINTS)


def brand_core_from_clean_text(text: str) -> str:
    """Extract brand label from a cleaned restaurant/branch name string."""
    text = _PAREN_SUFFIX_RE.sub(" ", text)
    core = text.split(",")[0].strip()
    if " - " in core:
        left, right = core.split(" - ", 1)
        if _looks_like_branch_location_suffix(right):
            core = left.strip()
    core = re.sub(r"\s+", " ", core).strip(" -_")
    return core


def normalize_restaurant_name(name: object) -> Optional[str]:
    """
    Brand-level name used for actual_unique entities.

    Strips delivery-area / branch location suffixes so that
    "CocoaVia, Hawally", "Asha's, Jabriya", "Shake Shack - 360 Mall, Zahra",
    and "BURGER BOUTIQUE" collapse to one restaurant each when they share a
    brand name across restaurantIds.
    """
    text = _clean_text(name)
    if text is None:
        return None
    core = brand_core_from_clean_text(text)
    if not core:
        return None
    return core.lower()


def restaurant_entity_key(row: dict[str, Any]) -> str:
    """
    Actual unique restaurant entity: prefer special franchise brand, else
    normalized name, else restaurantId fallback.
    """
    brand = row.get("special_franchise")
    if brand:
        return f"name:{str(brand).lower()}"
    norm = row.get("normalized_name") or normalize_restaurant_name(row.get("name"))
    if norm:
        return f"name:{norm}"
    rid = row.get("restaurantId")
    return f"restaurantId:{rid}"


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * r * math.asin(math.sqrt(min(1.0, a)))


def coord_in_kuwait(lat: float, lon: float) -> bool:
    return KUWAIT_LAT[0] <= lat <= KUWAIT_LAT[1] and KUWAIT_LON[0] <= lon <= KUWAIT_LON[1]


DEFAULT_AREA_MAP_FILENAME = "shoparea_and_id.xlsx"


def area_map_key_under_prefix(prefix: str, filename: str = DEFAULT_AREA_MAP_FILENAME) -> str:
    prefix = (prefix or "").strip()
    if prefix and not prefix.endswith("/"):
        prefix += "/"
    return f"{prefix}{filename}"


def load_area_map_from_excel(source: Any, label: str) -> dict[int, str]:
    df = pd.read_excel(source)
    if "area_id" not in df.columns or "area_name" not in df.columns:
        raise SystemExit(f"Expected columns area_id, area_name in {label}")
    out: dict[int, str] = {}
    for _, r in df.iterrows():
        try:
            out[int(r["area_id"])] = str(r["area_name"])
        except (TypeError, ValueError):
            continue
    if not out:
        raise SystemExit(f"No area rows loaded from {label}")
    print(f"[ok] area map: {label} ({len(out)} areas)")
    return out


def load_area_map_local(path: Path) -> dict[int, str]:
    if not path.exists():
        raise SystemExit(f"Area map not found: {path}")
    return load_area_map_from_excel(path, str(path))


def load_area_map_r2(client, bucket: str, key: str) -> dict[int, str]:
    from io import BytesIO

    from botocore.exceptions import ClientError

    try:
        obj = client.get_object(Bucket=bucket, Key=key)
    except ClientError as exc:
        raise SystemExit(
            f"Area map not found in R2 at s3://{bucket}/{key}\n"
            f"Upload shoparea_and_id.xlsx under the day prefix. Original error: {exc}"
        ) from exc
    raw = obj["Body"].read()
    return load_area_map_from_excel(BytesIO(raw), f"s3://{bucket}/{key}")


def load_records_local(local_dir: Path, key_contains: str, max_files: Optional[int]) -> tuple[list[dict], list[str]]:
    files = list(iter_local_files(local_dir, key_contains))
    if max_files is not None:
        files = files[:max_files]
    records: list[dict] = []
    sources: list[str] = []
    for source, path in files:
        raw = path.read_bytes()
        batch = load_json_records_from_bytes(raw)
        for item in batch:
            item = dict(item)
            item["_source_file"] = source
            records.append(item)
        sources.append(source)
        print(f"[ok] {path.name} ({len(batch)} rows)")
    return records, sources


def load_records_r2(
    client,
    bucket: str,
    prefix: str,
    key_contains: str,
    max_files: Optional[int],
) -> tuple[list[dict], list[str]]:
    keys = list_r2_keys(client, bucket, prefix, key_contains)
    if max_files is not None:
        keys = keys[:max_files]
    records: list[dict] = []
    sources: list[str] = []
    print(f"Found {len(keys)} matching objects in s3://{bucket}/{prefix}")
    for key in keys:
        obj = client.get_object(Bucket=bucket, Key=key)
        raw = obj["Body"].read()
        batch = load_json_records_from_bytes(raw)
        for item in batch:
            item = dict(item)
            item["_source_file"] = key
            records.append(item)
        sources.append(key)
        print(f"[ok] {key} ({len(batch)} rows)")
    return records, sources


def build_branch_frame(records: list[dict], area_map: dict[int, str]) -> pd.DataFrame:
    """Deduplicate delivery-area duplicates by branchId (keep first seen)."""
    by_branch: dict[Any, dict[str, Any]] = {}
    for item in records:
        bid = item.get("branchId")
        if bid is None:
            continue
        if bid in by_branch:
            continue
        brand = detect_special_franchise(item.get("name"), item.get("branchName"))
        shop_area = item.get("shopArea")
        try:
            shop_area_int = int(shop_area) if shop_area is not None and str(shop_area) != "" else None
        except (TypeError, ValueError):
            shop_area_int = None
        try:
            shop_city_int = int(item["shopCity"]) if item.get("shopCity") is not None else None
        except (TypeError, ValueError):
            shop_city_int = None
        try:
            vertical = int(item["verticalType"]) if item.get("verticalType") is not None else None
        except (TypeError, ValueError):
            vertical = None
        lat = item.get("latitude")
        lon = item.get("longitude")
        try:
            lat_f = float(lat) if lat is not None else None
            lon_f = float(lon) if lon is not None else None
        except (TypeError, ValueError):
            lat_f, lon_f = None, None

        norm_name = normalize_restaurant_name(item.get("name"))
        by_branch[bid] = {
            "branchId": bid,
            "restaurantId": item.get("restaurantId"),
            "name": item.get("name"),
            "branchName": item.get("branchName"),
            "normalized_name": norm_name,
            "shopCity": shop_city_int,
            "shopCity_label": SHOPCITY_LABELS.get(shop_city_int, f"shopCity_{shop_city_int}"),
            "shopArea": shop_area_int,
            "shopArea_name": area_map.get(shop_area_int, f"unknown_area_{shop_area}"),
            "latitude": lat_f,
            "longitude": lon_f,
            "cuisineString": item.get("cuisineString") or "",
            "verticalType": vertical,
            "verticalType_name": VERTICAL_TYPE_NAMES.get(vertical, f"verticalType_{vertical}"),
            "special_franchise": brand,
            "entity_key": None,
            "_source_file": item.get("_source_file"),
        }

    rows = list(by_branch.values())
    for row in rows:
        row["entity_key"] = restaurant_entity_key(row)
    return pd.DataFrame(rows)


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {path}")


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe = sheet_name[:31] or "Sheet1"
            df.to_excel(writer, sheet_name=safe, index=False)
    print(f"Wrote {path}")


def _json_list(values: list[Any]) -> str:
    """Serialize a Python list for Excel cells (JSON array)."""
    return json.dumps(values, ensure_ascii=False)


def _sorted_unique_ids(series: pd.Series) -> list[Any]:
    ids: list[Any] = []
    seen: set[Any] = set()
    for value in series.dropna().tolist():
        try:
            key: Any = int(value)
        except (TypeError, ValueError):
            key = value
        if key in seen:
            continue
        seen.add(key)
        ids.append(key)
    return sorted(ids, key=lambda x: (str(type(x)), str(x)))


def _sorted_unique_texts(series: pd.Series) -> list[str]:
    texts: list[str] = []
    seen: set[str] = set()
    for value in series.dropna().tolist():
        text = str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        texts.append(text)
    return sorted(texts, key=str.lower)


def _display_name_for_group(g: pd.DataFrame, franchise: Optional[str]) -> str:
    if franchise:
        return franchise
    if g["normalized_name"].notna().any():
        # Prefer canonical brand from normalization; title-case short brands only.
        norm = str(g["normalized_name"].dropna().iloc[0])
        if len(g["name"].dropna()):
            raw = _clean_text(g["name"].mode().iloc[0]) or ""
            core = brand_core_from_clean_text(raw)
            if core:
                return core
        return norm
    if len(g["name"].dropna()):
        raw = _clean_text(g["name"].mode().iloc[0]) or ""
        core = brand_core_from_clean_text(raw)
        if core:
            return core
    return str(g["restaurantId"].iloc[0])


def answer_q1(df: pd.DataFrame) -> dict[str, Any]:
    unique_by_id = int(df["restaurantId"].nunique())
    unique_entities = int(df["entity_key"].nunique())
    special = (
        df[df["special_franchise"].notna()]
        .groupby("special_franchise")
        .agg(
            branches=("branchId", "nunique"),
            unique_restaurantIds=("restaurantId", "nunique"),
        )
        .reset_index()
        .sort_values("special_franchise")
    )
    # Brands where one normalized name spans multiple restaurantIds (same problem as McCafe).
    multi_rows = []
    for entity_key, g in df.groupby("entity_key", dropna=False):
        restaurant_ids = _sorted_unique_ids(g["restaurantId"])
        n_ids = len(restaurant_ids)
        if n_ids <= 1:
            continue
        franchise = (
            str(g["special_franchise"].dropna().iloc[0])
            if g["special_franchise"].notna().any()
            else None
        )
        restaurant_names = _sorted_unique_texts(g["name"])
        multi_rows.append(
            {
                "entity_key": entity_key,
                "display_name": _display_name_for_group(g, franchise),
                "restaurant_name": _display_name_for_group(g, franchise),
                "restaurant_ids": _json_list(restaurant_ids),
                "restaurant_names": _json_list(restaurant_names),
                "branches": int(g["branchId"].nunique()),
                "unique_restaurantIds": n_ids,
            }
        )
    name_multi_id = pd.DataFrame(multi_rows)
    if len(name_multi_id):
        name_multi_id = name_multi_id.sort_values(
            ["unique_restaurantIds", "branches"], ascending=[False, False]
        )
    return {
        "question_number": 1,
        "question": "عدد المطاعم الكلي / Total restaurants",
        "explanation": (
            "Two totals are reported. unique_by_restaurant_id counts distinct restaurantId values. "
            "actual_unique_restaurant_entities counts distinct restaurants by normalized name "
            "(and id only as fallback when name is missing): location suffixes after a comma are "
            "stripped so Asha's / CocoaVia / Burger Boutique / McDonald's / McCafe / Starbucks "
            "and similar brands count as one restaurant even when each branch has its own restaurantId."
        ),
        "methodology": (
            "Deduplicate rows by branchId across delivery-area files. "
            "entity_key = special franchise brand (McDonald's/McCafe/Starbucks) when detected, "
            "else normalize(name) (apostrophes unified, DH Kitchen/TGO parentheticals removed, "
            "text before first comma kept, lowercased), else restaurantId."
        ),
        "answer": {
            "unique_by_restaurant_id": unique_by_id,
            "actual_unique_restaurant_entities": unique_entities,
            "difference_id_vs_name_entities": unique_by_id - unique_entities,
            "brands_with_multiple_restaurantIds": int(len(name_multi_id)),
            "special_franchises": special.to_dict(orient="records"),
            # Full list (not top-N) — used by Excel sheet multi_id_brands.
            "multi_id_brands": name_multi_id.to_dict(orient="records"),
            "unique_branches_context": int(df["branchId"].nunique()),
        },
        "generated_at": utc_now(),
    }


def answer_q2(df: pd.DataFrame) -> dict[str, Any]:
    total = int(df["branchId"].nunique())
    return {
        "question_number": 2,
        "question": "عدد الفروع الكلي / Total unique branches",
        "explanation": (
            "Total physical branches as unique branchId after removing duplicates that appear "
            "in multiple delivery-area MergedRestaurantsInfo files."
        ),
        "methodology": "Deduplicate source rows by branchId, then nunique(branchId).",
        "answer": {
            "total_unique_branch_ids": total,
            "rows_before_branch_dedupe_note": (
                "Same branch can appear in several area files; this count is after dedupe."
            ),
        },
        "generated_at": utc_now(),
    }


def _restaurants_by_id_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """One row per restaurantId with id + name."""
    rows = []
    for rid, g in df.groupby("restaurantId", dropna=False):
        name = str(g["name"].mode().iloc[0]) if len(g["name"].dropna()) else ""
        franchise = (
            str(g["special_franchise"].dropna().iloc[0])
            if g["special_franchise"].notna().any()
            else None
        )
        rows.append(
            {
                "restaurantId": rid,
                "name": name,
                "normalized_name": g["normalized_name"].dropna().iloc[0]
                if g["normalized_name"].notna().any()
                else None,
                "entity_key": g["entity_key"].iloc[0],
                "display_name": _display_name_for_group(g, franchise),
                "branch_count": int(g["branchId"].nunique()),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values(["name", "restaurantId"], ascending=[True, True]).reset_index(drop=True)


def _restaurants_actual_unique_sheet(per_entity: pd.DataFrame) -> pd.DataFrame:
    """One row per actual-unique restaurant entity (normalized name) with id(s) + name."""
    if per_entity.empty:
        return pd.DataFrame(
            columns=[
                "entity_key",
                "restaurant_name",
                "restaurant_ids",
                "restaurant_names",
                "unique_restaurantIds",
                "branch_count",
                "entity_type",
            ]
        )
    cols = [
        "entity_key",
        "restaurant_name",
        "restaurant_ids",
        "restaurant_names",
        "unique_restaurantIds",
        "branch_count",
        "branch_ids",
        "branch_names",
        "entity_type",
        "normalized_name",
        "franchise",
    ]
    present = [c for c in cols if c in per_entity.columns]
    out = per_entity[present].copy()
    return out.sort_values(
        ["restaurant_name", "entity_key"], ascending=[True, True]
    ).reset_index(drop=True)


def _branches_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """One row per unique branchId with id + name."""
    cols = [
        "branchId",
        "branchName",
        "name",
        "restaurantId",
        "normalized_name",
        "entity_key",
        "special_franchise",
        "shopCity",
        "shopCity_label",
        "shopArea",
        "shopArea_name",
    ]
    present = [c for c in cols if c in df.columns]
    out = df[present].drop_duplicates(subset=["branchId"]).copy()
    out = out.rename(columns={"name": "restaurant_name"})
    return out.sort_values(["branchName", "branchId"], ascending=[True, True]).reset_index(drop=True)


def sheet_q1(df: pd.DataFrame, per_entity: pd.DataFrame, q1: dict[str, Any]) -> dict[str, pd.DataFrame]:
    ans = q1["answer"]
    explanation = pd.DataFrame(
        [
            {
                "question_number": 1,
                "question": q1["question"],
                "explanation": q1["explanation"],
                "methodology": q1["methodology"],
                "unique_by_restaurant_id": ans["unique_by_restaurant_id"],
                "actual_unique_restaurant_entities": ans["actual_unique_restaurant_entities"],
                "difference_id_vs_name_entities": ans["difference_id_vs_name_entities"],
                "brands_with_multiple_restaurantIds": ans["brands_with_multiple_restaurantIds"],
                "unique_branches_context": ans["unique_branches_context"],
                "generated_at": q1["generated_at"],
            }
        ]
    )
    by_id = _restaurants_by_id_sheet(df)
    actual = _restaurants_actual_unique_sheet(per_entity)
    special = pd.DataFrame(ans.get("special_franchises") or [])
    multi = pd.DataFrame(ans.get("multi_id_brands") or [])
    return {
        "explanation": explanation,
        "restaurants_by_id": by_id,
        "restaurants_actual_unique": actual,
        "special_franchises": special if len(special) else pd.DataFrame(columns=["special_franchise"]),
        "multi_id_brands": multi if len(multi) else pd.DataFrame(columns=["entity_key"]),
    }


def sheet_q2(df: pd.DataFrame, q2: dict[str, Any]) -> dict[str, pd.DataFrame]:
    ans = q2["answer"]
    explanation = pd.DataFrame(
        [
            {
                "question_number": 2,
                "question": q2["question"],
                "explanation": q2["explanation"],
                "methodology": q2["methodology"],
                "total_unique_branch_ids": ans["total_unique_branch_ids"],
                "generated_at": q2["generated_at"],
            }
        ]
    )
    branches = _branches_sheet(df)
    return {
        "explanation": explanation,
        "branches": branches,
    }


def branches_per_entity(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for entity_key, g in df.groupby("entity_key", dropna=False):
        franchise = None
        if g["special_franchise"].notna().any():
            franchise = str(g["special_franchise"].dropna().iloc[0])
        if franchise:
            entity_type = "special_franchise"
        elif g["normalized_name"].notna().any():
            entity_type = "normalized_name"
        else:
            entity_type = "restaurantId"

        branch_rows = (
            g.drop_duplicates(subset=["branchId"])
            .sort_values(["branchId"], kind="mergesort")
            .reset_index(drop=True)
        )
        restaurant_ids = _sorted_unique_ids(g["restaurantId"])
        restaurant_names = _sorted_unique_texts(g["name"])
        branch_ids = _sorted_unique_ids(branch_rows["branchId"])
        branch_names = _sorted_unique_texts(
            branch_rows["branchName"].fillna(branch_rows["name"])
        )
        branches_detail = []
        for _, r in branch_rows.iterrows():
            try:
                bid = int(r["branchId"]) if pd.notna(r["branchId"]) else r["branchId"]
            except (TypeError, ValueError):
                bid = r["branchId"]
            try:
                rid = int(r["restaurantId"]) if pd.notna(r["restaurantId"]) else r["restaurantId"]
            except (TypeError, ValueError):
                rid = r["restaurantId"]
            bname = r["branchName"] if pd.notna(r.get("branchName")) else None
            rname = r["name"] if pd.notna(r.get("name")) else None
            branches_detail.append(
                {
                    "branchId": bid,
                    "branchName": None if bname is None else str(bname),
                    "restaurantId": rid,
                    "restaurantName": None if rname is None else str(rname),
                }
            )

        restaurant_name = _display_name_for_group(g, franchise)
        rows.append(
            {
                "entity_type": entity_type,
                "entity_key": entity_key,
                "restaurant_name": restaurant_name,
                "display_name": restaurant_name,  # alias used by older callers / sort keys
                "franchise": franchise,
                "normalized_name": g["normalized_name"].dropna().iloc[0]
                if g["normalized_name"].notna().any()
                else None,
                # Keep a representative single id for summary joins; Excel Q3 uses the list columns.
                "restaurantId": restaurant_ids[0] if restaurant_ids else None,
                "restaurant_ids": _json_list(restaurant_ids),
                "restaurant_names": _json_list(restaurant_names),
                "branch_ids": _json_list(branch_ids),
                "branch_names": _json_list(branch_names),
                "branches": _json_list(branches_detail),
                "branch_count": int(len(branch_ids)),
                "unique_branchIds": int(len(branch_ids)),
                "unique_restaurantIds": int(len(restaurant_ids)),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values(
        ["branch_count", "restaurant_name"], ascending=[False, True]
    ).reset_index(drop=True)


def sheet_q3(per_entity: pd.DataFrame) -> dict[str, pd.DataFrame]:
    dist = (
        per_entity["branch_count"]
        .value_counts()
        .sort_index()
        .rename_axis("branch_count")
        .reset_index(name="restaurant_entities")
    )
    explanation = pd.DataFrame(
        [
            {
                "question_number": 3,
                "question": "عدد الفروع لكل مطعم / Branches per restaurant",
                "explanation": (
                    "One row per actual-unique restaurant (normalized name). "
                    "restaurant_ids / restaurant_names / branch_ids / branch_names / branches "
                    "are JSON arrays so multi-id brands (McCafe, Starbucks, Asha's, etc.) keep "
                    "the full list of ids and names in one row instead of a single restaurantId."
                ),
                "row_count": int(len(per_entity)),
            }
        ]
    )
    detail_cols = [
        "restaurant_name",
        "restaurant_ids",
        "restaurant_names",
        "branch_count",
        "unique_restaurantIds",
        "unique_branchIds",
        "branch_ids",
        "branch_names",
        "branches",
        "entity_key",
        "entity_type",
        "franchise",
        "normalized_name",
    ]
    present = [c for c in detail_cols if c in per_entity.columns]
    detail = per_entity[present].copy()
    return {
        "explanation": explanation,
        "branches_per_restaurant": detail,
        "branch_count_distribution": dist,
    }


def answer_q5(per_entity: pd.DataFrame, df: pd.DataFrame) -> dict[str, Any]:
    avg_entities = float(per_entity["branch_count"].mean()) if len(per_entity) else 0.0
    avg_by_rid = float(df.groupby("restaurantId")["branchId"].nunique().mean()) if len(df) else 0.0
    return {
        "question_number": 5,
        "question": "متوسط الفروع لكل مطعم / Average branches per restaurant",
        "explanation": (
            "Primary average uses name-based restaurant entities (actual unique). "
            "A secondary average uses raw restaurantId only, which inflates brands that assign a "
            "different restaurantId per branch (McDonald's/McCafe/Starbucks and others)."
        ),
        "methodology": "mean(branch_count) over name-based entity table from Q3; also mean over restaurantId groups.",
        "answer": {
            "avg_branches_per_restaurant_entity": round(avg_entities, 4),
            "avg_branches_per_restaurantId_raw": round(avg_by_rid, 4),
            "restaurant_entities": int(len(per_entity)),
            "total_branches": int(df["branchId"].nunique()),
            "median_branches_per_restaurant_entity": float(per_entity["branch_count"].median())
            if len(per_entity)
            else 0.0,
        },
        "generated_at": utc_now(),
    }


def sheet_q4(per_entity: pd.DataFrame) -> dict[str, pd.DataFrame]:
    multi = per_entity[per_entity["branch_count"] > 1].copy()
    explanation = pd.DataFrame(
        [
            {
                "question_number": 4,
                "question": "المطاعم متعددة الفروع / Multi-branch restaurants",
                "explanation": (
                    "Lists restaurant entities with more than one branchId. "
                    "Same schema as Q3: restaurant_ids / restaurant_names / branch_ids / "
                    "branch_names / branches are JSON arrays with the full multi-id detail."
                ),
                "row_count": int(len(multi)),
            }
        ]
    )
    detail_cols = [
        "restaurant_name",
        "restaurant_ids",
        "restaurant_names",
        "branch_count",
        "unique_restaurantIds",
        "unique_branchIds",
        "branch_ids",
        "branch_names",
        "branches",
        "entity_key",
        "entity_type",
        "franchise",
        "normalized_name",
    ]
    present = [c for c in detail_cols if c in multi.columns]
    return {
        "explanation": explanation,
        "multi_branch_restaurants": multi[present].copy() if present else multi,
    }


def sheet_q6(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Branch counts and unique restaurant entities per city
    city_branch = (
        df.groupby(["shopCity", "shopCity_label"], dropna=False)
        .agg(
            unique_branches=("branchId", "nunique"),
            unique_restaurant_ids=("restaurantId", "nunique"),
            unique_restaurant_entities=("entity_key", "nunique"),
        )
        .reset_index()
        .sort_values("unique_branches", ascending=False)
    )
    explanation = pd.DataFrame(
        [
            {
                "question_number": 6,
                "question": "توزيع المطاعم حسب المدينة / Distribution by shopCity",
                "explanation": (
                    "Grouped by numeric shopCity. Labels are inferred from dominant Kuwait areas "
                    "(Capital, Hawalli, Farwaniya, Ahmadi, Jahra, Mubarak Al-Kabeer). "
                    "unique_restaurant_entities uses normalized restaurant name (actual unique)."
                ),
            }
        ]
    )
    return {"explanation": explanation, "by_shopCity": city_branch}


def sheet_q7(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    area = (
        df.groupby(["shopArea", "shopArea_name"], dropna=False)
        .agg(
            unique_branches=("branchId", "nunique"),
            unique_restaurant_ids=("restaurantId", "nunique"),
            unique_restaurant_entities=("entity_key", "nunique"),
        )
        .reset_index()
        .sort_values("unique_branches", ascending=False)
    )
    explanation = pd.DataFrame(
        [
            {
                "question_number": 7,
                "question": "توزيع الفروع حسب المنطقة / Branch distribution by shopArea",
                "explanation": (
                    "Grouped by shopArea. Area names come from shoparea_and_id.xlsx "
                    "(area_id -> area_name), loaded from R2 under the day prefix in CI."
                ),
            }
        ]
    )
    return {"explanation": explanation, "by_shopArea": area}


def _area_density_frame(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        coords = g.dropna(subset=["latitude", "longitude"]).copy()
        in_kw = coords[
            coords.apply(lambda r: coord_in_kuwait(r["latitude"], r["longitude"]), axis=1)
        ] if len(coords) else coords

        median_lat = float(in_kw["latitude"].median()) if len(in_kw) else None
        median_lon = float(in_kw["longitude"].median()) if len(in_kw) else None

        valid_mask = []
        for _, r in coords.iterrows():
            ok = coord_in_kuwait(r["latitude"], r["longitude"])
            if ok and median_lat is not None and median_lon is not None:
                dist = haversine_km(r["latitude"], r["longitude"], median_lat, median_lon)
                ok = dist <= COORD_OUTLIER_KM
            valid_mask.append(ok)
        valid = coords.loc[[i for i, ok in zip(coords.index, valid_mask) if ok]] if len(coords) else coords

        if len(valid) >= 2:
            # Approximate coverage km^2 from lat/lon span (rough, for relative density).
            lat_span = max(valid["latitude"].max() - valid["latitude"].min(), 1e-5)
            lon_span = max(valid["longitude"].max() - valid["longitude"].min(), 1e-5)
            mean_lat = float(valid["latitude"].mean())
            km_lat = lat_span * 111.0
            km_lon = lon_span * 111.0 * max(0.2, abs(math.cos(math.radians(mean_lat))))
            area_km2 = max(km_lat * km_lon, 0.01)
        elif len(valid) == 1:
            area_km2 = 0.5  # single-point placeholder so density is defined
        else:
            area_km2 = None

        branch_n = int(g["branchId"].nunique())
        entity_n = int(g["entity_key"].nunique())
        density_branches = (branch_n / area_km2) if area_km2 else None
        density_entities = (entity_n / area_km2) if area_km2 else None

        records.append(
            {
                "shopArea": area_id,
                "shopArea_name": area_name,
                "unique_branches": branch_n,
                "unique_restaurant_entities": entity_n,
                "coords_total": int(len(coords)),
                "coords_valid_geo": int(len(valid)),
                "coords_flagged_wrong": int(len(coords) - len(valid)),
                "median_latitude": median_lat,
                "median_longitude": median_lon,
                "approx_coverage_km2": round(area_km2, 4) if area_km2 is not None else None,
                "density_branches_per_km2": round(density_branches, 4) if density_branches is not None else None,
                "density_restaurant_entities_per_km2": round(density_entities, 4)
                if density_entities is not None
                else None,
                "density_rank_basis": "density_restaurant_entities_per_km2 using geo-validated coords",
            }
        )
    out = pd.DataFrame(records)
    if out.empty:
        return out
    return out.sort_values(
        ["density_restaurant_entities_per_km2", "unique_branches"],
        ascending=[False, False],
        na_position="last",
    ).reset_index(drop=True)


def sheet_q8(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    density = _area_density_frame(df)
    explanation = pd.DataFrame(
        [
            {
                "question_number": 8,
                "question": "كثافة المطاعم حسب المنطقة / Restaurant density by area",
                "explanation": (
                    "Density is unique restaurant entities (normalized name / actual unique) per "
                    "approximate km^2 of the area's geo-validated coordinates. Coordinates outside "
                    "Kuwait bounds or farther than "
                    f"{COORD_OUTLIER_KM} km from the area median lat/lon are flagged as likely wrong "
                    "and excluded from coverage/density math. shopArea names from R2 shoparea_and_id.xlsx."
                ),
            }
        ]
    )
    flagged = []
    for (area_id, area_name), g in df.groupby(["shopArea", "shopArea_name"], dropna=False):
        coords = g.dropna(subset=["latitude", "longitude"])
        if coords.empty:
            continue
        in_kw = coords[coords.apply(lambda r: coord_in_kuwait(r["latitude"], r["longitude"]), axis=1)]
        if in_kw.empty:
            med_lat = med_lon = None
        else:
            med_lat = float(in_kw["latitude"].median())
            med_lon = float(in_kw["longitude"].median())
        for _, r in coords.iterrows():
            reasons = []
            if not coord_in_kuwait(r["latitude"], r["longitude"]):
                reasons.append("outside_kuwait_bbox")
            elif med_lat is not None and med_lon is not None:
                dist = haversine_km(r["latitude"], r["longitude"], med_lat, med_lon)
                if dist > COORD_OUTLIER_KM:
                    reasons.append(f"outlier_vs_area_median_{dist:.1f}km")
            if reasons:
                flagged.append(
                    {
                        "shopArea": area_id,
                        "shopArea_name": area_name,
                        "branchId": r["branchId"],
                        "restaurantId": r["restaurantId"],
                        "name": r["name"],
                        "latitude": r["latitude"],
                        "longitude": r["longitude"],
                        "flag_reason": ";".join(reasons),
                    }
                )
    flagged_df = pd.DataFrame(flagged)
    return {
        "explanation": explanation,
        "density_by_area": density,
        "flagged_coordinates": flagged_df if len(flagged_df) else pd.DataFrame(columns=["flag_reason"]),
    }


def sheet_q9(density: pd.DataFrame) -> dict[str, pd.DataFrame]:
    usable = density.dropna(subset=["density_restaurant_entities_per_km2"]).copy()
    # Prefer areas with enough branches so single-point placeholders do not dominate extremes.
    ranked = usable[usable["unique_branches"] >= 3].copy()
    if ranked.empty:
        ranked = usable
    top = ranked.head(20).copy()
    bottom = ranked.sort_values("density_restaurant_entities_per_km2", ascending=True).head(20).copy()
    explanation = pd.DataFrame(
        [
            {
                "question_number": 9,
                "question": "المناطق الأعلى/الأقل كثافة / Highest and lowest density areas",
                "explanation": (
                    "Uses density_restaurant_entities_per_km2 from Q8 (shopArea + geo-validated "
                    "lat/lon). Ranking limited to areas with at least 3 branches when available, "
                    "to reduce noise from tiny/single-coordinate areas."
                ),
            }
        ]
    )
    return {
        "explanation": explanation,
        "highest_density": top,
        "lowest_density": bottom,
    }


def sheet_q10(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    by_vertical = (
        df.groupby(["verticalType", "verticalType_name"], dropna=False)
        .agg(
            unique_branches=("branchId", "nunique"),
            unique_restaurant_ids=("restaurantId", "nunique"),
            unique_restaurant_entities=("entity_key", "nunique"),
        )
        .reset_index()
        .sort_values("unique_branches", ascending=False)
    )

    # Split cuisineString on commas for multi-tag distribution
    cuisine_counter: Counter[str] = Counter()
    cuisine_branch: dict[str, set[Any]] = defaultdict(set)
    for _, r in df.iterrows():
        raw = str(r["cuisineString"] or "").strip()
        parts = [p.strip() for p in raw.split(",") if p.strip()] or ["(blank)"]
        for part in parts:
            cuisine_counter[part] += 1
            cuisine_branch[part].add(r["branchId"])
    cuisine_df = pd.DataFrame(
        [
            {
                "cuisine": c,
                "branch_mentions": n,
                "unique_branches": len(cuisine_branch[c]),
            }
            for c, n in cuisine_counter.most_common()
        ]
    )

    cross = (
        df.groupby(["verticalType", "verticalType_name", "cuisineString"], dropna=False)
        .agg(unique_branches=("branchId", "nunique"))
        .reset_index()
        .sort_values("unique_branches", ascending=False)
    )

    explanation = pd.DataFrame(
        [
            {
                "question_number": 10,
                "question": "توزيع المطاعم حسب نوع النشاط / Distribution by activity type",
                "explanation": (
                    "Primary split uses verticalType with platform meaning: "
                    "0 restaurants/food, 1 grocery/supermarket, 2 pharmacy/health, 3 flowers, "
                    "4 electronics, 5 pet shops, 6 cosmetics/beauty, 9 specialty stores. "
                    "cuisineString sheet splits comma-separated cuisine tags; cross sheet keeps "
                    "full cuisineString with verticalType."
                ),
            }
        ]
    )
    return {
        "explanation": explanation,
        "by_verticalType": by_vertical,
        "by_cuisine_tag": cuisine_df,
        "vertical_x_cuisineString": cross,
    }


def sheet_q11(df: pd.DataFrame, per_entity: pd.DataFrame) -> dict[str, pd.DataFrame]:
    branches = int(df["branchId"].nunique())
    entities = int(len(per_entity))
    raw_rids = int(df["restaurantId"].nunique())
    summary = pd.DataFrame(
        [
            {
                "metric": "unique_branches",
                "value": branches,
            },
            {
                "metric": "unique_restaurant_entities_brand_aware",
                "value": entities,
            },
            {
                "metric": "unique_restaurantIds_raw",
                "value": raw_rids,
            },
            {
                "metric": "branches_per_restaurant_entity",
                "value": round(branches / entities, 4) if entities else None,
            },
            {
                "metric": "restaurant_entities_per_branch",
                "value": round(entities / branches, 4) if branches else None,
            },
            {
                "metric": "pct_branches_that_are_multi_branch_entities",
                "value": round(
                    100.0
                    * per_entity.loc[per_entity["branch_count"] > 1, "branch_count"].sum()
                    / branches,
                    2,
                )
                if branches
                else None,
            },
            {
                "metric": "pct_restaurant_entities_multi_branch",
                "value": round(100.0 * (per_entity["branch_count"] > 1).mean(), 2) if entities else None,
            },
        ]
    )
    by_city = (
        df.groupby(["shopCity", "shopCity_label"], dropna=False)
        .agg(branches=("branchId", "nunique"), restaurant_entities=("entity_key", "nunique"))
        .reset_index()
    )
    by_city["branches_per_restaurant_entity"] = by_city.apply(
        lambda r: round(r["branches"] / r["restaurant_entities"], 4) if r["restaurant_entities"] else None,
        axis=1,
    )
    explanation = pd.DataFrame(
        [
            {
                "question_number": 11,
                "question": "نسبة المطاعم مقابل الفروع / Restaurants vs branches ratio",
                "explanation": (
                    "Compares name-based restaurant entities (actual unique) to unique branchIds. "
                    "branches_per_restaurant_entity > 1 means the average entity has multiple branches. "
                    "Also includes the raw restaurantId count for contrast."
                ),
            }
        ]
    )
    return {
        "explanation": explanation,
        "ratio_summary": summary,
        "ratio_by_shopCity": by_city.sort_values("branches", ascending=False),
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Market structure insights Q1-Q11")
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
    p.add_argument(
        "--area-map",
        type=Path,
        default=None,
        help="Local Excel mapping area_id -> area_name (local mode only; optional)",
    )
    p.add_argument(
        "--area-map-key",
        default=os.environ.get("AREA_MAP_R2_KEY", ""),
        help=(
            "R2 object key for shoparea_and_id.xlsx. "
            "Default: <prefix>/shoparea_and_id.xlsx"
        ),
    )
    p.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "outputs" / "market_structure",
        help="Local staging dir (gitignored). In CI this folder is uploaded as a GitHub Actions artifact.",
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
            raise SystemExit(
                f"Area map not found. Place {DEFAULT_AREA_MAP_FILENAME} next to local JSON "
                "or pass --area-map"
            )
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
    per_entity = branches_per_entity(df)
    out = args.out_dir
    out.mkdir(parents=True, exist_ok=True)

    q1 = answer_q1(df)
    q2 = answer_q2(df)
    q5 = answer_q5(per_entity, df)

    write_json(out / "q01_total_restaurants.json", q1)
    write_json(out / "q02_total_branches.json", q2)
    write_json(out / "q05_avg_branches_per_restaurant.json", q5)

    write_excel(out / "q01_total_restaurants.xlsx", sheet_q1(df, per_entity, q1))
    write_excel(out / "q02_total_branches.xlsx", sheet_q2(df, q2))
    write_excel(out / "q03_branches_per_restaurant.xlsx", sheet_q3(per_entity))
    write_excel(out / "q04_multi_branch_restaurants.xlsx", sheet_q4(per_entity))
    write_excel(out / "q06_restaurants_by_city.xlsx", sheet_q6(df))
    write_excel(out / "q07_branches_by_area.xlsx", sheet_q7(df))

    q8_sheets = sheet_q8(df)
    write_excel(out / "q08_restaurant_density_by_area.xlsx", q8_sheets)
    write_excel(out / "q09_highest_lowest_density_areas.xlsx", sheet_q9(q8_sheets["density_by_area"]))
    write_excel(out / "q10_by_activity_type.xlsx", sheet_q10(df))
    write_excel(out / "q11_restaurants_vs_branches_ratio.xlsx", sheet_q11(df, per_entity))

    summary = {
        "generated_at": utc_now(),
        "mode": mode,
        "source_files": sources,
        "raw_rows_loaded": len(records),
        "unique_branches_after_dedupe": int(df["branchId"].nunique()),
        "questions": {
            "json": [
                "q01_total_restaurants.json",
                "q02_total_branches.json",
                "q05_avg_branches_per_restaurant.json",
            ],
            "excel": [
                "q01_total_restaurants.xlsx",
                "q02_total_branches.xlsx",
                "q03_branches_per_restaurant.xlsx",
                "q04_multi_branch_restaurants.xlsx",
                "q06_restaurants_by_city.xlsx",
                "q07_branches_by_area.xlsx",
                "q08_restaurant_density_by_area.xlsx",
                "q09_highest_lowest_density_areas.xlsx",
                "q10_by_activity_type.xlsx",
                "q11_restaurants_vs_branches_ratio.xlsx",
            ],
        },
        "headline": {
            "unique_by_restaurant_id": q1["answer"]["unique_by_restaurant_id"],
            "actual_unique_restaurant_entities": q1["answer"]["actual_unique_restaurant_entities"],
            "total_unique_branches": q2["answer"]["total_unique_branch_ids"],
            "avg_branches_per_restaurant_entity": q5["answer"]["avg_branches_per_restaurant_entity"],
        },
        "area_map_source": area_map_source,
        "reference_inputs": [
            "data/insightspart1.txt",
            area_map_source,
            "data/franchise_restaurant_id_check.json",
            "data/franchise_insights.json",
        ],
    }
    write_json(out / "summary.json", summary)
    print("")
    print(
        f"Done: branches={summary['unique_branches_after_dedupe']}, "
        f"entities={summary['headline']['actual_unique_restaurant_entities']}, "
        f"files={len(sources)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
