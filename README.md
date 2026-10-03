# TBDA-TA

Talabat branch / restaurant data analysis helpers.

## Merged restaurant attribute profiler

Scans merged-restaurant JSON files under this R2 prefix only:

```text
merged-restaurant-info/year=2025/month=09/day=17/
```

That includes every area folder underneath, for example:

```text
merged-restaurant-info/year=2025/month=09/day=17/10-kaifan/10-kaifan_MergedRestaurantsInfo.json
merged-restaurant-info/year=2025/month=09/day=17/100-mina-abdullah/100-mina-abdullah_MergedRestaurantsInfo.json
```

Report includes meaning, observed type, example values, and Filled/Empty/Missing %.

### Outputs (artifacts only — not committed)

`outputs/` is gitignored. In GitHub Actions, download artifact `restaurant-attribute-profile`:

- `important_attributes_with_examples.txt`
- `attribute_profile_summary.json`

### Required secrets

- `CF_R2_ACCESS_KEY_ID`
- `CF_R2_SECRET_ACCESS_KEY`
- `CF_R2_ENDPOINT_URL`
- `CF_R2_BUCKET_NAME`

### GitHub Actions

Workflow: `.github/workflows/profile_restaurant_attributes.yml`

- Default prefix: `merged-restaurant-info/year=2025/month=09/day=17/`
- Default key filter: `MergedRestaurantsInfo.json`
- Uploads artifact only (does not commit outputs)

### Local run against R2

```bash
export CF_R2_ACCESS_KEY_ID=...
export CF_R2_SECRET_ACCESS_KEY=...
export CF_R2_ENDPOINT_URL=...
export CF_R2_BUCKET_NAME=...
export R2_PREFIX=merged-restaurant-info/year=2025/month=09/day=17/

python scripts/profile_merged_restaurant_attributes.py --mode r2
```
