#!/usr/bin/env python3
"""
Unique restaurants by `id` only (no name merging).

Reads all merged-restaurant-info JSON from Cloudflare R2 (CI) or local files,
keeps one row per distinct `id`, and writes Excel + summary under --out-dir.

In CI the out-dir is uploaded as a GitHub Actions artifact only
(not committed to the repo, not written back to R2).

Columns:
  id, name, logo, restaurantSlug, branchId, branchName, branchSlug,
  cuisineString, shopArea, area_name, shopCity, isGrocery, verticalType,
  shopType, isDarkstore

`branchId` / `branchName` / `branchSlug` are JSON lists of all distinct branches
for that restaurant id (one or many elements). Other fields use the first-seen row.
`area_name` comes from shoparea_and_id.xlsx by matching shopArea -> area_id.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from io import BytesIO
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

DEFAULT_AREA_MAP_FILENAME = "shoparea_and_id.xlsx"

COLUMNS: list[str] = [
    "id",
    "name",
    "logo",
    "restaurantSlug",
    "branchId",
    "branchName",
    "branchSlug",
    "cuisineString",
    "shopArea",
    "area_name",
    "shopCity",
    "isGrocery",
    "verticalType",
    "shopType",
    "isDarkstore",
]


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def resolve_restaurant_id(item: dict[str, Any]) -> Any:
    """Prefer `id`; fall back to `restaurantId` (they are equal in this dataset)."""
    rid = item.get("id")
    if rid is None or rid == "":
        rid = item.get("restaurantId")
    return rid


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
    from botocore.exceptions import ClientError

    try:
        obj = client.get_object(Bucket=bucket, Key=key)
    except ClientError as exc:
        raise SystemExit(
            f"Area map not found in R2 at s3://{bucket}/{key}\n"
            f"Upload shoparea_and_id.xlsx under the day prefix. Original error: {exc}"
        ) from exc
    return load_area_map_from_excel(BytesIO(obj["Body"].read()), f"s3://{bucket}/{key}")


def shop_area_int(value: Any) -> Optional[int]:
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def area_name_for(shop_area: Any, area_map: dict[int, str]) -> Optional[str]:
    aid = shop_area_int(shop_area)
    if aid is None:
        return None
    return area_map.get(aid)


def load_records_local(
    local_dir: Path, key_contains: str, max_files: Optional[int]
) -> tuple[list[dict], list[str]]:
    files = list(iter_local_files(local_dir, key_contains))
    if max_files is not None:
        files = files[:max_files]
    records: list[dict] = []
    sources: list[str] = []
    for source, path in files:
        batch = load_json_records_from_bytes(path.read_bytes())
        for item in batch:
            row = dict(item)
            row["_source_file"] = source
            records.append(row)
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
        batch = load_json_records_from_bytes(obj["Body"].read())
        for item in batch:
            row = dict(item)
            row["_source_file"] = key
            records.append(row)
        sources.append(key)
        print(f"[ok] {key} ({len(batch)} rows)")
    return records, sources


BRANCH_LIST_COLUMNS: tuple[str, ...] = ("branchId", "branchName", "branchSlug")


def format_list_for_excel(values: list[Any]) -> str:
    """Always a JSON list string (one or many elements) so Excel keeps full branch info."""
    return json.dumps(values, ensure_ascii=False)


def extract_restaurant_fields(
    item: dict[str, Any], restaurant_id: Any, area_map: dict[int, str]
) -> dict[str, Any]:
    """Restaurant-level fields (first-seen). Branch columns are filled later as lists."""
    shop_area = item.get("shopArea")
    return {
        "id": restaurant_id,
        "name": item.get("name"),
        "logo": item.get("logo"),
        "restaurantSlug": item.get("restaurantSlug"),
        "cuisineString": item.get("cuisineString"),
        "shopArea": shop_area,
        "area_name": area_name_for(shop_area, area_map),
        "shopCity": item.get("shopCity"),
        "isGrocery": item.get("isGrocery"),
        "verticalType": item.get("verticalType"),
        "shopType": item.get("shopType"),
        "isDarkstore": item.get("isDarkstore"),
    }


def unique_by_id(
    records: list[dict], area_map: dict[int, str]
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    One row per distinct restaurant `id`.

    `branchId` / `branchName` / `branchSlug` are lists of all distinct branches for
    that id (deduped by branchId). Restaurant-level fields use the first-seen row.
    `area_name` is joined from shoparea_and_id.xlsx via shopArea == area_id.
    """
    by_id: dict[Any, dict[str, Any]] = {}
    # rid -> ordered branchId -> {branchId, branchName, branchSlug}
    branches_by_id: dict[Any, dict[Any, dict[str, Any]]] = {}
    skipped_missing_id = 0

    for item in records:
        rid = resolve_restaurant_id(item)
        if rid is None or rid == "":
            skipped_missing_id += 1
            continue

        if rid not in by_id:
            by_id[rid] = extract_restaurant_fields(item, rid, area_map)
            branches_by_id[rid] = {}

        bid = item.get("branchId")
        if bid is None or bid == "":
            continue
        if bid in branches_by_id[rid]:
            continue
        branches_by_id[rid][bid] = {
            "branchId": bid,
            "branchName": item.get("branchName"),
            "branchSlug": item.get("branchSlug"),
        }

    rows: list[dict[str, Any]] = []
    for rid in sorted(by_id.keys(), key=lambda x: (str(type(x)), x)):
        row = dict(by_id[rid])
        branch_rows = list(branches_by_id.get(rid, {}).values())
        row["branchId"] = [b["branchId"] for b in branch_rows]
        row["branchName"] = [b["branchName"] for b in branch_rows]
        row["branchSlug"] = [b["branchSlug"] for b in branch_rows]
        rows.append(row)

    df = pd.DataFrame(rows, columns=COLUMNS)
    df = df.sort_values(
        by="name",
        ascending=True,
        kind="mergesort",
        key=lambda s: s.astype(str).str.strip().str.casefold(),
    ).reset_index(drop=True)

    branch_counts = df["branchId"].map(len)
    multi_branch_ids = int((branch_counts > 1).sum()) if len(df) else 0
    matched = int(df["area_name"].notna().sum()) if len(df) else 0
    stats = {
        "raw_rows_loaded": len(records),
        "skipped_missing_id": skipped_missing_id,
        "unique_restaurant_ids": int(len(df)),
        "ids_with_multiple_branches": multi_branch_ids,
        "unique_branchIds_across_restaurants": int(branch_counts.sum()) if len(df) else 0,
        "area_name_matched": matched,
        "area_name_unmatched": int(len(df) - matched),
        "area_map_size": len(area_map),
        "branch_list_columns": list(BRANCH_LIST_COLUMNS),
    }
    return df, stats


def dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Serialize list columns to JSON strings for openpyxl (always list form)."""
    out = df.copy()
    for col in BRANCH_LIST_COLUMNS:
        out[col] = out[col].map(lambda v: format_list_for_excel(list(v) if v is not None else []))
    return out


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Unique restaurants by id only → Excel artifact")
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
        help="Local Excel mapping area_id -> area_name (local mode; optional auto-detect)",
    )
    p.add_argument(
        "--area-map-key",
        default=os.environ.get("AREA_MAP_R2_KEY", ""),
        help="R2 object key for shoparea_and_id.xlsx. Default: <prefix>/shoparea_and_id.xlsx",
    )
    p.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "outputs" / "unique_restaurants_by_id",
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

    df, stats = unique_by_id(records, area_map)
    out = args.out_dir
    out.mkdir(parents=True, exist_ok=True)

    excel_name = "unique_restaurants_by_id.xlsx"
    excel_path = out / excel_name
    excel_df = dataframe_for_excel(df)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        excel_df.to_excel(writer, sheet_name="restaurants", index=False)
        pd.DataFrame(
            [
                {
                    "note": (
                        "One row per distinct restaurant `id` across all MergedRestaurantsInfo "
                        "files. branchId / branchName / branchSlug are JSON lists of every distinct "
                        "branch for that id (always a list, even when length is 1). "
                        "Delivery-area duplicates of the same branchId are collapsed. "
                        "Restaurant-level fields (name, logo, shopArea, etc.) come from the "
                        "first-seen row. area_name is joined from shoparea_and_id.xlsx where "
                        "shopArea == area_id."
                    )
                }
            ]
        ).to_excel(writer, sheet_name="readme", index=False)

    summary = {
        "generated_at": utc_now(),
        "mode": mode,
        "dedupe_key": "id",
        "columns": COLUMNS,
        "source_files": sources,
        "source_file_count": len(sources),
        "area_map_source": area_map_source,
        **stats,
        "outputs": {
            "excel": excel_name,
            "summary": "summary.json",
        },
        "delivery": (
            "GitHub Actions artifact only — not committed to the repo and not uploaded to R2."
        ),
    }
    write_json(out / "summary.json", summary)

    print("")
    print(
        f"Done: unique_ids={stats['unique_restaurant_ids']}, "
        f"multi_branch={stats['ids_with_multiple_branches']}, "
        f"branches={stats['unique_branchIds_across_restaurants']}, "
        f"area_name_matched={stats['area_name_matched']}, "
        f"raw_rows={stats['raw_rows_loaded']}, files={len(sources)}"
    )
    print(f"Excel: {excel_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
