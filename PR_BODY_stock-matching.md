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

- [x] I ran the fixture preview (`apps/dashboards/scripts/preview_local.sh` equivalent —
      real `uvicorn server.main:app` with `DASHBOARDS_DEV=1 DASHBOARDS_FIXTURE=1`, fixtures
      auto-staged into `data/` per the registry, on 2026-07-10) and reviewed the rendered
      dashboard end-to-end, including a real bug found and fixed during that review.

```
Stock Matching — fixture preview verified 2026-07-10
  Server:  uvicorn server.main:app, DASHBOARDS_DEV=1 DASHBOARDS_FIXTURE=1, port 8080
  Registry: python -c "from server.registry import DASHBOARDS; ..." → stock-matching
            correctly discovered as data_mode=frozen-snapshot, alongside rtb-ab / ai-readiness
  Page:    GET /d/stock-matching/ → 200; GET /d/stock-matching/data/meta.js → 200,
           fixture content served (labelled "FIXTURE DATA — not real (synthetic)")
  Rendered (fixture values): Suppliers 3 · Total 6,000 · Mapped 5,100 (85.0%) · Not mapped 900 (15.0%)
  Chart: inline SVG stacked bars, 3 suppliers, sorted by total — matches KPI totals
  Table: 3 supplier rows sum to the All-suppliers total (6,000/5,100/900/15.0%);
         click-to-expand verified on demo-alpha → Demo Alpha NSW (2,200/2,000/200) +
         Demo Alpha VIC (1,100/1,000/100), sums back to the parent row exactly
  Methodology/caveats section (Chapel Corner note, snapshot-not-live, aggregate-only) renders correctly

BUG FOUND + FIXED during this review (commit 19645d8, already pushed to this branch):
  .kpi .label only set color, not background. The shared platform theme
  (web/assets/popeye-theme.css) defines a bare .label as a badge component
  (white text on a #a1a1a1 grey pill) — its background leaked through onto
  every KPI tile, rendering the "Suppliers"/"Total in-stock rows"/"Mapped"/
  "Not mapped" tile labels almost invisible (grey-on-grey, confirmed via
  computed styles: color rgb(154,163,172) vs background-color rgb(161,161,161)).
  This would have shipped to production, not just the fixture preview.
  Fix: renamed the class to .kpi-label (CSS + the renderKpis() template
  string) so it no longer collides with the shared badge class. Re-verified
  visually after the fix — background-color now rgba(0,0,0,0), label text
  legible in the intended muted-secondary colour.

Validation:
  python -m server.registry --validate  → OK (stock-matching discovered, 11 dashboards)
  scripts/verify_manifest.py             → stock-matching data/ matches MANIFEST (meta.js 466B, suppliers.js 18,229B)
```
