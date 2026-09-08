# Vehicle fitment dashboard

Static dashboard showing how many logged vehicles at Tyroola have tyre sizes on
file and how many have none, for AU and NZ. No build step, no dependencies —
open `index.html` in a browser.

```
index.html    markup only
styles.css    all styling
app.js        chart rendering (monthly chart, donut drill-down, trend chart)
data.js       every figure on the page, as window.VLDATA
queries/      the BigQuery SQL each figure comes from
vehicle_log_ev_report.xlsx   the row-level report the page links to
```

## The headline

Of 322,433 vehicles in the log, **72,465 (22.5%) have tyre sizes on file and
249,968 (77.5%) have none.**

The 249,968 split three ways, and the split matters because each has a
different owner:

| State | Vehicles | What it means |
|---|---|---|
| Matched, no sizes | 146,865 | The spec exists in the fitment table but carries no sizes. Catalogue work. |
| No matching spec | 98,597 | Make/model/series resolves to nothing. Mapping and naming work. |
| No make recorded | 4,506 | Nothing was captured to look up. Capture work in the lookup flow. |

## Data sources

All in BigQuery project `heroic-ruler-198603`:

- `popeye_mirror.vehicle_log` — one row per logged vehicle lookup
- `RTB.vehicle_catalog_ev_ready` — make/model to BEV/PHEV/HEV/ICE, 1,191 pairs
- `tyroola_reporting_prod.vehicle_tyre_fitment` — single snapshot, 2026-06-16

Log coverage: 2021-11-22 to 2026-09-01.

## Two joins that need explaining

**Fitment joins on make + model + series_chassis, never on vehicle_id.** The
fitment table's own schema describes its `vehicle_id` as "session-ephemeral; NOT
a stable identity key". Testing confirms it: of 50,811 ids that match the log,
99.6% agree on make and model but the year almost never does, so the id resolves
to a different variant. Only 128 of those 50,811 came back fitted — 0.25%
against a 34.8% base rate. If a future change makes that join look tempting,
this is why it was rejected.

**EV class is resolved in priority order**: engine string containing ELEC or BEV
wins first, then the catalogue's `ev_vehicle` flag, then blank engine with no
catalogue match becomes Unknown. HEV and PHEV count as EV, following the
catalogue's own flag.

## Known data problems

These are surfaced on the page itself, and are the reason no figure here should
be quoted without a caveat:

1. **`tyre_options` is the literal string `[]` on all 324,350 rows.** So
   `request_count` and `request_count_without_tyre_options` are identical
   throughout. 1,140 rows carry a real `tyre_options_added_at` timestamp anyway,
   which suggests the field was populated upstream and lost in the mirror.
2. **May and June 2026 hold 123,590 vehicles**, 38% of the whole log, against a
   normal 5,000–7,000 a month. Almost all unfitted. Looks like a bulk load, and
   it drags every whole-log rate down. Confirm what it was before quoting the
   headline.
3. **`updated_at` precedes `created_at` on 313,070 of 324,350 rows.** Dates here
   take the later of the two; per-request timing can't be trusted until this is
   explained.
4. **The AU_SCA iframe channel stopped writing on 2025-12-08.** It is the source
   of the 4,506 no-make records; they carry 29,540 requests, averaging 6.6 each
   against about 1.0 elsewhere.
5. **The fitment snapshot is three months behind the log.** Vehicles first seen
   after 2026-06-16 can count as unmatched for no other reason.
6. **Model strings fragment.** Grouping is on the raw string, so Ford Falcon
   holds 4,161 vehicles under `FALCON` and 8 more under seven spellings,
   including two of "FALCON GL,S-PAK" that differ by one space.
7. **`fitted` means listed, not served.** A vehicle counts as fitted when its
   spec resolves to a fitment row with sizes. Because `tyre_options` is empty,
   the log cannot confirm what the customer was actually shown.

Cross-project source `api-project-290933994367` is not reachable from this
project, so none of the above could be checked upstream.

## Refreshing the data

`data.js` is a plain object, hand-refreshable. Run the queries in `queries/`
against BigQuery and update the matching keys:

- `months[]` — from `02_monthly_series.sql`; keys are `m`, `w` (with sizes),
  `o` (without), then `e` / `n` / `u` for unfitted EV / Non EV / Unknown
- `types[]` and `lastRequestedYear[]` — from `03_type_and_year_rollups.sql`
- `hero[]` and `meta` — totals; `hero[]` must sum to `meta.vehicles_total`

Sanity checks worth keeping: `hero[]` sums to the total, and for each row in
`types[]`, `has + nosizes + nospec + nomake == vehicles`.

## Notes for future changes

The chart frame (`.plot`, `.ylab`, `.parea`, `.grid`, `.chart`, `.ticks`) is
shared by all three charts. `drawAxis()` and `drawTicks()` in `app.js` are the
shared helpers — change the axis behaviour there once rather than per chart.

Both time charts clip outlier months rather than rescaling: the axis is set to
2.5× the median of the visible months, and taller bars are drawn hatched with
their true total printed above. This exists because of problem 2 above. The
monthly chart offers a full-scale toggle if you want the honest magnitude.
