#!/usr/bin/env python3
"""
Unique restaurants by `id` only (no name merging).

Reads all merged-restaurant-info JSON from Cloudflare R2 (CI) or local files,
keeps one row per distinct `id`, and writes Excel + summary under --out-dir.

In CI the out-dir is uploaded as a GitHub Actions artifact only
(not committed to the repo, not written back to R2).

Columns:
  id, name, restaurantSlug, branchId, branchName, branchSlug,
  cuisineString, shopArea, shopCity, isGrocery, verticalType,
  shopType, isDarkstore
"""

from __future__ import annotations

import argparse
import json
import os
import sys
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

COLUMNS: list[str] = [
    "id",
    "name",
    "restaurantSlug",
    "branchId",
    "branchName",
    "branchSlug",
    "cuisineString",
    "shopArea",
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


def extract_row(item: dict[str, Any], restaurant_id: Any) -> dict[str, Any]:
    return {
        "id": restaurant_id,
        "name": item.get("name"),
        "restaurantSlug": item.get("restaurantSlug"),
        "branchId": item.get("branchId"),
        "branchName": item.get("branchName"),
        "branchSlug": item.get("branchSlug"),
        "cuisineString": item.get("cuisineString"),
        "shopArea": item.get("shopArea"),
        "shopCity": item.get("shopCity"),
        "isGrocery": item.get("isGrocery"),
        "verticalType": item.get("verticalType"),
        "shopType": item.get("shopType"),
        "isDarkstore": item.get("isDarkstore"),
    }


def unique_by_id(records: list[dict]) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Keep first-seen row per distinct restaurant `id`.

    Same id can appear across many delivery-area files and (for multi-branch
    brands) with different branchId values; uniqueness is id-only as requested.
    """
    by_id: dict[Any, dict[str, Any]] = {}
    branch_ids_per_id: dict[Any, set[Any]] = {}
    skipped_missing_id = 0

    for item in records:
        rid = resolve_restaurant_id(item)
        if rid is None or rid == "":
            skipped_missing_id += 1
            continue

        bid = item.get("branchId")
        if rid not in branch_ids_per_id:
            branch_ids_per_id[rid] = set()
        if bid is not None and bid != "":
            branch_ids_per_id[rid].add(bid)

        if rid in by_id:
            continue
        by_id[rid] = extract_row(item, rid)

    rows = [by_id[k] for k in sorted(by_id.keys(), key=lambda x: (str(type(x)), x))]
    df = pd.DataFrame(rows, columns=COLUMNS)

    multi_branch_ids = sum(1 for s in branch_ids_per_id.values() if len(s) > 1)
    stats = {
        "raw_rows_loaded": len(records),
        "skipped_missing_id": skipped_missing_id,
        "unique_restaurant_ids": int(len(df)),
        "ids_with_multiple_branchIds_before_dedupe": multi_branch_ids,
        "unique_branchIds_seen_across_raw": int(
            len({b for s in branch_ids_per_id.values() for b in s})
        ),
    }
    return df, stats


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
        "--out-dir",
        type=Path,
        default=ROOT / "outputs" / "unique_restaurants_by_id",
        help="Local staging dir (gitignored). In CI uploaded as a GitHub Actions artifact.",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    if args.mode == "local":
        local_dir = args.local_dir
        if local_dir is None:
            candidate = ROOT.parent / "Restaurant_info"
            local_dir = candidate if candidate.exists() else ROOT / "sample_data"
        if not local_dir.exists():
            raise SystemExit(f"Local directory not found: {local_dir}")
        records, sources = load_records_local(local_dir, args.key_contains, args.max_files)
        mode = "local"
    else:
        bucket = clean_env("CF_R2_BUCKET_NAME")
        client = make_r2_client(bucket=bucket)
        records, sources = load_records_r2(
            client, bucket, args.prefix, args.key_contains, args.max_files
        )
        mode = "r2"

    if not records:
        raise SystemExit("No restaurant records loaded")

    df, stats = unique_by_id(records)
    out = args.out_dir
    out.mkdir(parents=True, exist_ok=True)

    excel_name = "unique_restaurants_by_id.xlsx"
    excel_path = out / excel_name
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="restaurants", index=False)
        pd.DataFrame(
            [
                {
                    "note": (
                        "One row per distinct restaurant `id` across all MergedRestaurantsInfo "
                        "files. Duplicate delivery-area listings and extra branches for the same "
                        "id are collapsed; branch* / shopArea / shopCity come from the first-seen row."
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
        f"raw_rows={stats['raw_rows_loaded']}, files={len(sources)}"
    )
    print(f"Excel: {excel_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
