#!/usr/bin/env python3
"""
Profile merged-restaurant-info attributes across Cloudflare R2 (or local files).

Writes local reports under --out-dir (default: outputs/, gitignored).
In CI these are uploaded as GitHub Actions artifacts only (not committed).

Default R2 prefix (day folder only):
  merged-restaurant-info/year=2025/month=09/day=17/
Layout under that prefix:
  <area>/<area>_MergedRestaurantsInfo.json
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator, Optional

# Allow `python scripts/...` from repo root
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.attribute_catalog import SECTIONS, UNIQUE_ATTRIBUTES, AttributeSpec  # noqa: E402


EMPTY_MARKERS = (None, "", [], {})


@dataclass
class PathStats:
    present: int = 0
    filled: int = 0
    empty: int = 0
    missing: int = 0
    types: Counter = field(default_factory=Counter)
    examples: list[Any] = field(default_factory=list)
    example_sources: list[str] = field(default_factory=list)

    def observe(self, value: Any, source: str, max_examples: int = 3) -> None:
        self.present += 1
        if _is_empty(value):
            self.empty += 1
        else:
            self.filled += 1
            self.types[_type_name(value)] += 1
            if len(self.examples) < max_examples and not _example_already_seen(self.examples, value):
                self.examples.append(_compact_example(value))
                self.example_sources.append(source)


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    if value == "":
        return True
    if value == []:
        return True
    if value == {}:
        return True
    return False


def _type_name(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int):
        return "int"
    if isinstance(value, float):
        return "float"
    if isinstance(value, str):
        return "string"
    if isinstance(value, list):
        if not value:
            return "array"
        first = value[0]
        if isinstance(first, dict):
            return "array<object>"
        if isinstance(first, bool):
            return "array<bool>"
        if isinstance(first, int):
            return "array<int>"
        if isinstance(first, float):
            return "array<float>"
        if isinstance(first, str):
            return "array<string>"
        return "array"
    if isinstance(value, dict):
        return "object"
    return type(value).__name__


def _compact_example(value: Any, max_str: int = 120) -> Any:
    if isinstance(value, str):
        return value if len(value) <= max_str else value[: max_str - 3] + "..."
    if isinstance(value, list):
        if not value:
            return []
        head = value[0]
        if isinstance(head, dict):
            return [_compact_example(head)]
        preview = value[:3]
        return preview + (["..."] if len(value) > 3 else [])
    if isinstance(value, dict):
        out = {}
        for i, (k, v) in enumerate(value.items()):
            if i >= 6:
                out["..."] = f"+{len(value) - 6} more"
                break
            out[k] = _compact_example(v)
        return out
    return value


def _example_already_seen(examples: list[Any], value: Any) -> bool:
    compact = _compact_example(value)
    return any(json.dumps(e, sort_keys=True, default=str) == json.dumps(compact, sort_keys=True, default=str) for e in examples)


def _get_by_path(record: dict, path: str) -> tuple[bool, Any]:
    """
    Resolve dotted paths, including array parents:
      cuisines.id -> values from each cuisine object
      availablePaymentMethods.dimension.width -> values from each method
    Returns (found_any_parent_context, value_or_list_of_values).
    For nested array paths, value is a list of leaf values (may be empty if parent arrays empty).
    """
    parts = path.split(".")
    # Fast path: top-level
    if len(parts) == 1:
        if parts[0] in record:
            return True, record[parts[0]]
        return False, None

    # Walk, expanding lists when encountered
    current: list[Any] = [record]
    for part in parts:
        next_level: list[Any] = []
        found_here = False
        for node in current:
            if isinstance(node, dict) and part in node:
                found_here = True
                val = node[part]
                if isinstance(val, list):
                    next_level.extend(val)
                else:
                    next_level.append(val)
            elif isinstance(node, list):
                for item in node:
                    if isinstance(item, dict) and part in item:
                        found_here = True
                        val = item[part]
                        if isinstance(val, list):
                            next_level.extend(val)
                        else:
                            next_level.append(val)
        if not found_here and not next_level:
            return False, None
        current = next_level

    # Nested path result is always observed as individual leaf values
    return True, current


def _nested_units(record: dict, path: str) -> list[Any]:
    """Return leaf observations for a path (0..N values)."""
    found, value = _get_by_path(record, path)
    if not found:
        return []  # missing
    if "." not in path:
        return [value]
    # value is list of leaves for nested paths
    if isinstance(value, list):
        return value
    return [value]


@dataclass
class RunSummary:
    mode: str
    files_listed: int = 0
    files_parsed: int = 0
    files_failed: int = 0
    restaurants: int = 0
    failed_keys: list[str] = field(default_factory=list)
    source_files: list[str] = field(default_factory=list)


def iter_local_files(local_dir: Path, key_contains: str) -> Iterator[tuple[str, Path]]:
    for path in sorted(local_dir.rglob("*.json")):
        if key_contains.lower() in path.name.lower():
            yield str(path), path


def make_r2_client():
    import boto3
    from botocore.config import Config

    access_key = os.environ.get("CF_R2_ACCESS_KEY_ID")
    secret_key = os.environ.get("CF_R2_SECRET_ACCESS_KEY")
    endpoint = os.environ.get("CF_R2_ENDPOINT_URL")
    if not all([access_key, secret_key, endpoint]):
        raise SystemExit(
            "Missing R2 credentials. Set CF_R2_ACCESS_KEY_ID, "
            "CF_R2_SECRET_ACCESS_KEY, CF_R2_ENDPOINT_URL."
        )
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=os.environ.get("AWS_REGION", "auto"),
        config=Config(signature_version="s3v4", retries={"max_attempts": 10, "mode": "standard"}),
    )


def _key_matches(key: str, key_contains: str) -> bool:
    """Match hive-style merged restaurant keys."""
    name = key.split("/")[-1]
    if not name.lower().endswith(".json"):
        return False
    needle = (key_contains or "").lower()
    if not needle:
        return name.lower().endswith("mergedrestaurantsinfo.json")
    return needle in name.lower() or needle in key.lower()


def list_r2_keys(client, bucket: str, prefix: str, key_contains: str) -> list[str]:
    keys: list[str] = []
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix or ""):
        for obj in page.get("Contents", []) or []:
            key = obj["Key"]
            if _key_matches(key, key_contains):
                keys.append(key)
    return sorted(keys)


def load_json_records_from_bytes(raw: bytes) -> list[dict]:
    data = json.loads(raw.decode("utf-8"))
    if isinstance(data, list):
        return [r for r in data if isinstance(r, dict)]
    if isinstance(data, dict):
        # common wrappers
        for k in ("restaurants", "data", "items", "results"):
            if isinstance(data.get(k), list):
                return [r for r in data[k] if isinstance(r, dict)]
        return [data]
    return []


def profile_records(
    records: Iterable[dict],
    stats: dict[str, PathStats],
    summary: RunSummary,
    source: str,
    max_examples: int,
) -> None:
    for record in records:
        summary.restaurants += 1
        for spec in UNIQUE_ATTRIBUTES:
            path = spec.path
            if "." not in path:
                if path not in record:
                    stats[path].missing += 1
                    continue
                stats[path].observe(record[path], source, max_examples=max_examples)
            else:
                leaves = _nested_units(record, path)
                parent_top = path.split(".")[0]
                if parent_top not in record:
                    stats[path].missing += 1
                    continue
                parent_val = record[parent_top]
                if _is_empty(parent_val):
                    # parent present but empty -> count one empty observation
                    stats[path].observe(parent_val if parent_val is not None else None, source, max_examples=max_examples)
                    continue
                if not leaves:
                    stats[path].empty += 1
                    stats[path].present += 1
                    continue
                for leaf in leaves:
                    stats[path].observe(leaf, source, max_examples=max_examples)


def run_local(local_dir: Path, key_contains: str, max_examples: int) -> tuple[dict[str, PathStats], RunSummary]:
    stats = {a.path: PathStats() for a in UNIQUE_ATTRIBUTES}
    summary = RunSummary(mode="local")
    files = list(iter_local_files(local_dir, key_contains))
    summary.files_listed = len(files)
    for key, path in files:
        try:
            raw = path.read_bytes()
            records = load_json_records_from_bytes(raw)
            profile_records(records, stats, summary, source=path.name, max_examples=max_examples)
            summary.files_parsed += 1
            summary.source_files.append(str(path))
            print(f"[ok] {path.name} ({len(records)} restaurants)")
        except Exception as exc:  # noqa: BLE001
            summary.files_failed += 1
            summary.failed_keys.append(f"{key}: {exc}")
            print(f"[fail] {path}: {exc}", file=sys.stderr)
    return stats, summary


def run_r2(prefix: str, key_contains: str, max_examples: int, max_files: Optional[int]) -> tuple[dict[str, PathStats], RunSummary]:
    bucket = os.environ.get("CF_R2_BUCKET_NAME")
    if not bucket:
        raise SystemExit("Missing CF_R2_BUCKET_NAME")

    client = make_r2_client()
    keys = list_r2_keys(client, bucket, prefix, key_contains)
    if max_files is not None:
        keys = keys[:max_files]

    stats = {a.path: PathStats() for a in UNIQUE_ATTRIBUTES}
    summary = RunSummary(mode="r2")
    summary.files_listed = len(keys)
    print(f"Found {len(keys)} matching objects in s3://{bucket}/{prefix}")

    for key in keys:
        try:
            obj = client.get_object(Bucket=bucket, Key=key)
            raw = obj["Body"].read()
            records = load_json_records_from_bytes(raw)
            profile_records(records, stats, summary, source=key, max_examples=max_examples)
            summary.files_parsed += 1
            summary.source_files.append(key)
            print(f"[ok] {key} ({len(records)} restaurants)")
        except Exception as exc:  # noqa: BLE001
            summary.files_failed += 1
            summary.failed_keys.append(f"{key}: {exc}")
            print(f"[fail] {key}: {exc}", file=sys.stderr)
    return stats, summary


def pct(n: int, d: int) -> float:
    return 0.0 if d <= 0 else (100.0 * n / d)


def format_example(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, (dict, list)):
        text = json.dumps(value, ensure_ascii=False, separators=(", ", ": "))
    else:
        text = str(value)
    text = text.replace("\n", " ")
    if len(text) > 140:
        return text[:137] + "..."
    return text


def _basename(source: str) -> str:
    return source.replace("\\", "/").split("/")[-1]


def build_text_report(stats: dict[str, PathStats], summary: RunSummary) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    lines: list[str] = []
    lines.append("merged-restaurant-info")
    lines.append("Important Attributes with Examples, Meaning, and Fill Rates")
    lines.append("")
    lines.append(f"Generated           : {now}")
    lines.append(f"Mode                : {summary.mode}")
    lines.append(f"Files listed        : {summary.files_listed}")
    lines.append(f"Files parsed        : {summary.files_parsed}")
    lines.append(f"Files failed        : {summary.files_failed}")
    lines.append(f"Restaurants scanned : {summary.restaurants}")
    lines.append("")
    lines.append("Definitions")
    lines.append("  Filled %       = non-null / non-empty observations")
    lines.append("  Empty/Null %   = null, '', [], or {} observations")
    lines.append("  Missing %      = attribute absent from the restaurant record")
    lines.append("  Nested fields  = rates over nested items when parent objects/arrays exist")
    lines.append("")

    by_section: dict[str, list[AttributeSpec]] = defaultdict(list)
    seen_paths: set[str] = set()
    for spec in UNIQUE_ATTRIBUTES:
        if spec.path in seen_paths:
            continue
        seen_paths.add(spec.path)
        by_section[spec.section].append(spec)

    for section in SECTIONS:
        lines.append("=" * 88)
        lines.append(section)
        lines.append("=" * 88)
        for spec in by_section.get(section, []):
            st = stats[spec.path]
            if "." not in spec.path:
                denom = summary.restaurants
                missing = st.missing
                if denom > 0 and (st.present + missing) < denom:
                    missing = denom - st.present
            else:
                denom = st.present + st.missing
                missing = st.missing

            filled_p = pct(st.filled, denom if denom else 1)
            empty_p = pct(st.empty, denom if denom else 1)
            missing_p = pct(missing, denom if denom else 1)
            observed_type = st.types.most_common(1)[0][0] if st.types else spec.expected_type
            example = format_example(st.examples[0]) if st.examples else "(no filled example found)"
            source = _basename(st.example_sources[0]) if st.example_sources else "-"

            indent = "  " * spec.nested_level
            lines.append(f"{indent}{spec.path}")
            lines.append(f"{indent}  Meaning      : {spec.meaning}")
            lines.append(f"{indent}  Expected type: {spec.expected_type}")
            lines.append(f"{indent}  Observed type: {observed_type}")
            lines.append(f"{indent}  Filled %     : {filled_p:6.2f}%   ({st.filled}/{denom if denom else 0})")
            lines.append(f"{indent}  Empty/Null % : {empty_p:6.2f}%   ({st.empty}/{denom if denom else 0})")
            lines.append(f"{indent}  Missing %    : {missing_p:6.2f}%   ({missing}/{denom if denom else 0})")
            lines.append(f"{indent}  Example      : {example}")
            lines.append(f"{indent}  Example from : {source}")
            lines.append("")
        lines.append("")

    lines.append("=" * 88)
    lines.append("Notes")
    lines.append("=" * 88)
    lines.append("cuisines[]                     nested cuisine objects (id, name, slug)")
    lines.append("Sponsored{}                    sponsored placement object (category, type, token)")
    lines.append("availablePaymentMethods[]      payment method objects (logo, text, dimension.width/height)")
    if summary.failed_keys:
        lines.append("")
        lines.append("Failed files")
        lines.append("-" * 88)
        for item in summary.failed_keys[:50]:
            lines.append(f"- {item}")
        if len(summary.failed_keys) > 50:
            lines.append(f"... +{len(summary.failed_keys) - 50} more")
    lines.append("")
    return "\n".join(lines)


def build_json_summary(stats: dict[str, PathStats], summary: RunSummary) -> dict:
    attributes = []
    for spec in UNIQUE_ATTRIBUTES:
        st = stats[spec.path]
        if "." not in spec.path:
            denom = summary.restaurants or 1
            missing = st.missing if (st.present + st.missing) >= summary.restaurants else max(summary.restaurants - st.present, 0)
        else:
            denom = (st.present + st.missing) or 1
            missing = st.missing
        attributes.append(
            {
                "path": spec.path,
                "section": spec.section,
                "meaning": spec.meaning,
                "expected_type": spec.expected_type,
                "observed_types": dict(st.types),
                "filled": st.filled,
                "empty_or_null": st.empty,
                "missing": missing,
                "present": st.present,
                "denominator": denom if summary.restaurants else 0,
                "filled_pct": round(pct(st.filled, denom), 4),
                "empty_or_null_pct": round(pct(st.empty, denom), 4),
                "missing_pct": round(pct(missing, denom), 4),
                "examples": st.examples,
                "example_sources": st.example_sources,
            }
        )
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "mode": summary.mode,
        "files_listed": summary.files_listed,
        "files_parsed": summary.files_parsed,
        "files_failed": summary.files_failed,
        "restaurants_scanned": summary.restaurants,
        "source_files": summary.source_files,
        "failed_files": summary.failed_keys,
        "attributes": attributes,
    }


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Profile merged-restaurant-info attributes from R2 or local files")
    p.add_argument("--mode", choices=["r2", "local"], default=os.environ.get("PROFILE_MODE", "r2"))
    p.add_argument("--local-dir", type=Path, default=None, help="Local directory of JSON files (local mode)")
    p.add_argument(
        "--prefix",
        default=os.environ.get(
            "R2_PREFIX",
            "merged-restaurant-info/year=2025/month=09/day=17/",
        ),
        help="R2 key prefix (default: merged-restaurant-info/year=2025/month=09/day=17/)",
    )
    p.add_argument(
        "--key-contains",
        default=os.environ.get("R2_KEY_CONTAINS", "MergedRestaurantsInfo.json"),
        help="Only include keys/filenames containing this substring",
    )
    p.add_argument("--max-files", type=int, default=None, help="Optional cap on number of files")
    p.add_argument("--max-examples", type=int, default=3)
    p.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "outputs",
        help="Output directory for report files",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()
    args.out_dir.mkdir(parents=True, exist_ok=True)

    if args.mode == "local":
        local_dir = args.local_dir
        if local_dir is None:
            # convenient default when running beside the downloaded dataset
            candidate = ROOT.parent / "Restaurant_info"
            local_dir = candidate if candidate.exists() else ROOT / "sample_data"
        if not local_dir.exists():
            raise SystemExit(f"Local directory not found: {local_dir}")
        stats, summary = run_local(local_dir, args.key_contains, args.max_examples)
    else:
        stats, summary = run_r2(args.prefix, args.key_contains, args.max_examples, args.max_files)

    text_path = args.out_dir / "important_attributes_with_examples.txt"
    json_path = args.out_dir / "attribute_profile_summary.json"
    text_path.write_text(build_text_report(stats, summary), encoding="utf-8")
    json_path.write_text(json.dumps(build_json_summary(stats, summary), indent=2, ensure_ascii=False), encoding="utf-8")

    print("")
    print(f"Wrote {text_path}")
    print(f"Wrote {json_path}")
    print(
        f"Done: parsed={summary.files_parsed}/{summary.files_listed}, "
        f"restaurants={summary.restaurants}, failed={summary.files_failed}"
    )
    return 0 if summary.files_failed == 0 or summary.files_parsed > 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
