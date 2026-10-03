#!/usr/bin/env python3
"""Quick R2 permission diagnostic for GitHub Actions / local use."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "profile_merged_restaurant_attributes.py"


def _load_profile_module():
    spec = importlib.util.spec_from_file_location("profile_merged_restaurant_attributes", MODULE_PATH)
    if spec is None or spec.loader is None:
        raise SystemExit(f"Cannot load {MODULE_PATH}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main() -> int:
    mod = _load_profile_module()
    bucket = mod._clean_env("CF_R2_BUCKET_NAME")
    prefix = mod._clean_env("R2_PREFIX") or "merged-restaurant-info/year=2025/month=09/day=17/"
    client = mod.make_r2_client(bucket=bucket)
    mod.probe_r2_access(client, bucket, prefix)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
