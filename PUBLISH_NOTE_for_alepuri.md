# Publish request — Supplier Stock Mapping dashboard

**From:** Emeterio (emeterio@tyroola.com)
**For:** alepuri (dashboards site owner)

## What I'd like published

A new dashboard tab on the internal dashboards site:
https://dashboards-vuweqgdcca-ts.a.run.app/

- **Tab name:** Stock Matching
- **Section:** Marketing (same group as the existing tabs, e.g. Shipment Accuracy)
- **Suggested anchor:** `#stock-matching`
- **File to publish:** `supplier_stock_match_dashboard_inlined.html` (recommended — see below)

## About the file

- **Single, fully self-contained HTML file** — no build step, and **zero external
  dependencies**. Data and the charting library are both inlined directly in the file;
  there is no backend call, query, or CDN fetch at runtime. It will render on any
  network / CSP policy.
- **Title:** "Supplier Stock Mapping Report" — mapped vs. not-mapped live stock rows by supplier, with a per-location drill-down.
- Chart.js v4.4.1 is inlined into the file, so no `script-src` / CSP changes are needed.

## Two versions are provided (use the inlined one)

- `supplier_stock_match_dashboard_inlined.html` — **use this.** Chart.js inlined, no CDN.
- `supplier_stock_match_dashboard.html` — original; loads Chart.js from
  `cdnjs.cloudflare.com`. Only use this if you'd rather allow the CDN in the site CSP.

## Data provenance (for reference)

- Project: `heroic-ruler-198603`
- Source table: `xx_development.popeye_production_au_public_supplier_stock`
- Joined to: `Sales_Dashboard_Views.suppliers_check_dashboard_table` for supplier keys
- Filters: `qty > 0`, `supplier_pid NOT LIKE 'DELETED\_%'`
- Note: this source is a static snapshot, not a live feed — rates are directional.

## Snapshot of the numbers (as built)

- 31 suppliers, 101,957 in-stock rows
- 90,232 mapped (88.5%) / 11,725 not mapped (11.5%)

## Request: repo access for next time

Could you also send me the **repo URL** and add me as a **collaborator**
(my GitHub username is **Metstyroola**)? That way I can push future dashboards
straight to the right folder/branch myself instead of routing each one through you.

Thanks!
