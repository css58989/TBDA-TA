"""Load merged-restaurant-info records from Cloudflare R2 or local JSON files."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterator, Optional


def clean_env(name: str) -> str:
    value = os.environ.get(name, "")
    return value.strip().strip('"').strip("'")


def normalize_endpoint(endpoint: str, bucket: str) -> str:
    from urllib.parse import urlparse

    endpoint = endpoint.strip().rstrip("/")
    if not endpoint.startswith("http://") and not endpoint.startswith("https://"):
        endpoint = "https://" + endpoint
    if bucket and endpoint.endswith("/" + bucket):
        endpoint = endpoint[: -(len(bucket) + 1)]

    parsed = urlparse(endpoint)
    host = parsed.netloc.lower()
    if bucket and host.startswith(bucket.lower() + ".") and host.endswith(".r2.cloudflarestorage.com"):
        account_host = host[len(bucket) + 1 :]
        endpoint = f"{parsed.scheme}://{account_host}"
        print(f"Normalized endpoint host from bucket subdomain -> {endpoint}")
    return endpoint


def make_r2_client(bucket: Optional[str] = None):
    import boto3
    from botocore.config import Config

    access_key = clean_env("CF_R2_ACCESS_KEY_ID")
    secret_key = clean_env("CF_R2_SECRET_ACCESS_KEY")
    endpoint = clean_env("CF_R2_ENDPOINT_URL")
    bucket = (bucket or clean_env("CF_R2_BUCKET_NAME")).strip()
    if not all([access_key, secret_key, endpoint, bucket]):
        raise SystemExit(
            "Missing R2 credentials. Set CF_R2_ACCESS_KEY_ID, "
            "CF_R2_SECRET_ACCESS_KEY, CF_R2_ENDPOINT_URL, CF_R2_BUCKET_NAME."
        )

    endpoint = normalize_endpoint(endpoint, bucket)
    print(f"R2 endpoint : {endpoint}")
    print(f"R2 bucket   : {bucket}")
    print(f"R2 key id   : {access_key[:4]}...{access_key[-4:] if len(access_key) > 8 else ''}")

    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=clean_env("AWS_REGION") or "auto",
        config=Config(
            signature_version="s3v4",
            s3={"addressing_style": "path"},
            retries={"max_attempts": 10, "mode": "standard"},
        ),
    )


def key_matches(key: str, key_contains: str) -> bool:
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
            if key_matches(key, key_contains):
                keys.append(key)
    return sorted(keys)


def iter_local_files(local_dir: Path, key_contains: str) -> Iterator[tuple[str, Path]]:
    for path in sorted(local_dir.rglob("*.json")):
        if key_matches(str(path), key_contains):
            yield str(path), path


def load_json_records_from_bytes(raw: bytes) -> list[dict]:
    import json

    data = json.loads(raw.decode("utf-8"))
    if isinstance(data, list):
        return [r for r in data if isinstance(r, dict)]
    if isinstance(data, dict):
        for k in ("restaurants", "data", "items", "results"):
            if isinstance(data.get(k), list):
                return [r for r in data[k] if isinstance(r, dict)]
        return [data]
    return []
