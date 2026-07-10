# Publish request (LIVE version) — Supplier Stock Mapping dashboard

**From:** Emeterio (emeterio@tyroola.com)
**For:** alepuri (dashboards site owner)

This is a follow-up to the static version. Same tab ("Stock Matching", Marketing
section) — but this one queries BigQuery so it stays current instead of being a
frozen snapshot.

## What "live" needs

The dashboard reads its data from an endpoint (`GET /api/stock-mapping`) that runs a
BigQuery query and returns JSON. For that to work on the dashboards Cloud Run site,
**two things are required** that only you can set up:

1. **A backend route** on the dashboards service that runs the query and returns the
   JSON (contract below). The site currently serves static files, so this is the new bit.
2. **BigQuery access for the service's runtime service account** in project
   `heroic-ruler-198603`:
   - `roles/bigquery.jobUser` (to run queries), and
   - `roles/bigquery.dataViewer` on dataset `xx_development` (source table).

If standing up a backend isn't worth it, see **Option B** — a scheduled refresh that
keeps it static-but-current with no runtime backend.

## Files provided

- `stock_mapping_query.sql` — the exact, self-contained query (grouping baked in via a
  mapping CTE). Verified: 31 supplier keys, 90,232 matched / 13,498 not-mapped.
- `Stock_Matcher/supplier_stock_match_dashboard_live.html` — the dashboard page; on load
  it `fetch('/api/stock-mapping')` and renders. (Chart.js still from CDN here; inline it
  if the site CSP blocks CDNs — same note as the static version.)
- `Stock_Matcher/server.js` — a **working reference implementation** of the endpoint
  (see the `/api/stock-mapping` route and `fetchStockMapping()` / `STOCK_MAP_NAME_TO_KEY`).
  It already runs this end-to-end locally.

## Endpoint contract — `GET /api/stock-mapping`

```jsonc
{
  "ok": true,
  "generated_at": "2026-07-10T01:04:34.309Z",  // when the query ran
  "refreshed_at": "2026-07-01 06:24:45+00",     // MAX(refreshed_at) of the source table
  "suppliers": 31,
  "totals": { "matched": 90232, "unmatched": 13498, "total": 103730 },
  "data": [
    {
      "key": "tempetyres",
      "matched": 28306, "unmatched": 2214, "total": 30520,
      "locations": [
        { "name": "Tempetyres VIC", "matched": 6869, "unmatched": 702, "total": 7571 }
        // ...
      ]
    }
    // ... 31 suppliers, sorted by total desc
  ]
}
```

The `data` array is exactly the shape the dashboard expects (it replaces the old inline
`var data = [...]`). Roll the SQL's location-level rows up by `supplier_key`; keep the
per-location rows under `locations`.

## Option A — true live (backend route)

Wire a route on the dashboards service that runs `stock_mapping_query.sql`, reshapes the
rows into the contract above, and returns JSON. `server.js` does exactly this in ~40 lines
if you want to lift it. Grant the service account the two BQ roles listed above.

## Option B — scheduled "static but current" (no backend)

If you'd rather not run a backend: a scheduled job (Cloud Scheduler → tiny Cloud Function,
or a cron in your build) runs `stock_mapping_query.sql` every N hours, writes the `data`
array into the HTML (or a `stock_mapping.json` sidecar the page fetches), and republishes.
Near-live without changing how the site serves content.

## Curation notes / caveats

- Source table `xx_development.popeye_production_au_public_supplier_stock` is a **periodic
  refresh, not a real-time feed** — `refreshed_at` tells you how fresh it is.
- Filters: `qty > 0`, `supplier_pid NOT LIKE 'DELETED\_%'`, restricted to the curated
  31-supplier account set (30 low-priority/duplicate accounts — Web Dummy, Tyres4u,
  promo duplicates, etc. — are excluded via the allow-list in the query).
- **Chapel Corner:** the original static snapshot had non-tyre accessory SKUs (valves,
  wheel weights, sockets) removed via a manual SKU-level list I don't have. The live query
  does **not** remove them, so Chapel Corner's not-mapped count runs ~1,773 higher than the
  snapshot. If you have/obtain that exclusion list, add it to the query's WHERE clause.

Thanks!
