## Summary

Adds a new **Stock Matching** dashboard under the **Marketing** category.

Supplier stock-mapping coverage — mapped vs not-mapped in-stock rows by supplier,
with a per-location drill-down. New bundle under the analyst-writable path
`apps/dashboards/web/dashboards/stock-matching/` (registry auto-discovers it; no
platform-owner files touched).

- `dashboard.json` — `category: "Marketing"`, `data_mode: "frozen-snapshot"`, `data_freeze: "2026-07-01"`
- `index.html` — house `popeye-theme.css`, inline SVG stacked bars (no CDN), click-to-expand
  location detail; loads `data/*.js` with a synthetic `fixtures/*.js` fallback + no-data banner
- `MANIFEST.md` — source, filters, Chapel Corner caveat, approved-file hashes
- `fixtures/` — synthetic placeholder data for local rendering

**Data source:** frozen aggregate snapshot from
`heroic-ruler-198603.xx_development.popeye_production_au_public_supplier_stock`
(`match_status`, `qty>0`, exclude `DELETED` pids, curated 31-supplier account set).
Aggregate only — supplier/location row counts, no SKU/order/customer data.

> **Operational step (before merge/preview passes):** the real snapshot files
> (`meta.js`, `suppliers.js`) must be staged to `gs://betty-boop-data/stock-matching/`
> so the build's `verify_manifest` finds them (they're gitignored, per the platform's
> data rules). Hashes are pinned in `MANIFEST.md`. This is a bucket-access action
> (@alepuri or an analyst with access) — the files are provided alongside this PR.

## Dashboard preview evidence

### Dashboard UI change (`apps/dashboards/web/dashboards/**`)

- [x] I reviewed the dashboard rendered locally (served the platform `web/` dir with the
      real snapshot staged into `data/`).

```
Stock Matching — local render verified 2026-07-10
  Category: Marketing · data_mode: frozen-snapshot (freeze 2026-07-01)
  KPIs:  Suppliers 31 · Total in-stock rows 103,730 · Mapped 90,232 (87.0%) · Not mapped 13,498 (13.0%)
  Chart: inline SVG stacked bars — 31 suppliers × mapped/not-mapped = 62 segments, sorted by total
  Table: 31 supplier rows + 117 location rows (click-to-expand); All-suppliers total 103,730 / 90,232 / 13,498 / 13.0%
  No-data path: shows "snapshot data missing" banner when data/ is absent
Validation:
  python -m server.registry --validate  → OK (stock-matching discovered, 11 dashboards)
  scripts/verify_manifest.py             → stock-matching data/ matches MANIFEST (meta.js 466B, suppliers.js 18,229B)
```
