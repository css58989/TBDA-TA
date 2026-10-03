#!/usr/bin/env python3
"""Quick R2 permission diagnostic for GitHub Actions / local use."""

from __future__ import annotations

import os
import sys
from urllib.parse import urlparse

import boto3
from botocore.config import Config
from botocore.exceptions import ClientError


def clean(name: str) -> str:
    return os.environ.get(name, "").strip().strip('"').strip("'")


def normalize_endpoint(endpoint: str, bucket: str) -> str:
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


def main() -> int:
    access_key = clean("CF_R2_ACCESS_KEY_ID")
    secret_key = clean("CF_R2_SECRET_ACCESS_KEY")
    endpoint = clean("CF_R2_ENDPOINT_URL")
    bucket = clean("CF_R2_BUCKET_NAME")
    prefix = clean("R2_PREFIX") or "merged-restaurant-info/year=2025/month=09/day=17/"

    if not all([access_key, secret_key, endpoint, bucket]):
        print("Missing one or more R2 secrets", file=sys.stderr)
        return 1

    endpoint = normalize_endpoint(endpoint, bucket)
    print(f"R2 endpoint : {endpoint}")
    print(f"R2 bucket   : {bucket}")
    print(f"R2 key id   : {access_key[:4]}...{access_key[-4:]}")
    print(f"R2 key len  : access={len(access_key)} secret={len(secret_key)}")
    print(f"R2 prefix   : {prefix}")

    client = boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name=clean("AWS_REGION") or "auto",
        config=Config(
            signature_version="s3v4",
            s3={"addressing_style": "path"},
            retries={"max_attempts": 5, "mode": "standard"},
        ),
    )

    checks = [
        ("HeadBucket", lambda: client.head_bucket(Bucket=bucket)),
        ("ListObjectsV2(root, MaxKeys=1)", lambda: client.list_objects_v2(Bucket=bucket, MaxKeys=1)),
        (
            f"ListObjectsV2(prefix={prefix!r}, MaxKeys=1)",
            lambda: client.list_objects_v2(Bucket=bucket, Prefix=prefix, MaxKeys=1),
        ),
    ]

    print("R2 permission probes:")
    failed = 0
    for name, fn in checks:
        try:
            fn()
            print(f"  [ok]   {name}")
        except ClientError as exc:
            code = exc.response.get("Error", {}).get("Code", "Unknown")
            print(f"  [fail] {name} -> {code}: {exc}")
            failed += 1
        except Exception as exc:  # noqa: BLE001
            print(f"  [fail] {name} -> {exc}")
            failed += 1

    if failed:
        print(
            "\nAccess denied usually means the R2 API token is scoped to a different "
            "bucket than CF_R2_BUCKET_NAME, or lacks List permission.\n"
            "Create a new token: R2 -> Manage R2 API Tokens -> Admin Read "
            "(or Object Read) applied to this exact bucket."
        )
        return 1
    print("\nAll probes passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
