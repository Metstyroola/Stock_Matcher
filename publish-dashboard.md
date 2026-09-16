# Publish the description-churn dashboard

Rebuild the mapped-SKU description churn dashboard from BigQuery and publish it to GitHub Pages.

Run this when someone asks to refresh, rebuild, or publish the dashboard. It is safe to re-run:
every step overwrites its own output.

---

## Config

```
GITHUB_OWNER   Metstyroola
GITHUB_REPO    <SET THIS — the repo to publish from>
BRANCH         main
PAGES_SOURCE   GitHub Actions
BQ_PROJECT     heroic-ruler-198603
BQ_TABLE       heroic-ruler-198603.popeye_mirror.supplier_description_change_log
YEAR_START     2026-01-01
```

`GITHUB_REPO` is the only value that must be filled in before the first run. If the repo does not
exist, create it (private is fine — Pages works on private repos for GitHub Pro/Team/Enterprise; on
a free account the repo must be public for Pages to serve). Ask before creating a repo or making one
public — do not decide either on the user's behalf.

## Repo layout

```
.github/workflows/deploy.yml   Pages deploy, triggered on push to docs/
scripts/extract.sh             BigQuery -> data/raw_edits.json
scripts/build.py               raw_edits.json -> docs/index.html (+ embedded xlsx)
templates/dashboard.html       HTML shell with __DATA__ and __XLSX__ placeholders
data/raw_edits.json            extract output (committed, so builds are reproducible)
docs/index.html                the published artifact
publish-dashboard.md           this file
```

---

## Steps

### 1. Extract from BigQuery

```bash
bash scripts/extract.sh
```

Writes `data/raw_edits.json`. Read the row-count reconciliation it prints before continuing.

**The query tool caps at 3,000 rows per call and truncates silently.** `extract.sh` pulls in
month-sized chunks and splits any month that comes back at exactly 3,000. It then re-queries a
single aggregate `COUNT(*)` and compares. If the counts disagree, stop and narrow the chunks —
do not build on a short extract. This has already bitten once: a 16–31 May window came back at
exactly 3,000 rows and silently dropped the last five days of the month.

If `bq` is unavailable, run the same SQL through the BigQuery MCP connector and write the rows to
`data/raw_edits.json` in the same shape (`{"cols": [...], "rows": [[...], ...]}`).

### 2. Build

```bash
python3 scripts/build.py
```

Reads `data/raw_edits.json` and `templates/dashboard.html`, writes `docs/index.html` with the
aggregates and the xlsx embedded as base64. Prints the numbers it derived — check them against
step 1 before publishing. It also prints the date range, which month (if any) is drawn as partial,
and the average line, all derived from the data; the page renders those from the payload rather
than from text in the template.

Needs `openpyxl`. Install with `pip install openpyxl` if missing.

### 3. Verify before publishing

Do all four. A broken dashboard published is worse than a stale one.

- `docs/index.html` is roughly 0.9–1.1 MB. Much smaller means the xlsx did not embed.
- The embedded base64 decodes to a valid workbook:
  ```bash
  python3 - <<'EOF'
  import re, base64, io
  from openpyxl import load_workbook
  h = open('docs/index.html').read()
  b64 = re.search(r'<script id="xlsx" type="text/plain">(.*?)</script>', h, re.S).group(1).strip()
  wb = load_workbook(io.BytesIO(base64.b64decode(b64)))
  print(wb.sheetnames, wb['Edits'].max_row - 4, 'rows')
  EOF
  ```
- Open `docs/index.html` in a browser. Both charts render, the download button produces a file.
- The SKU total in the header matches what `build.py` printed.

### 4. Publish

Commit `data/raw_edits.json` and `docs/index.html` together, push to `BRANCH`. The Actions workflow
deploys `docs/` to Pages.

**Ask the user before pushing.** Show them the numbers from step 2 and the diff stat first. Pushing
is what makes it live, so it is their call, not yours — including on a re-run.

Commit message form:

```
Refresh churn dashboard through <last changed_at date>

<N> edits, <M> SKUs, <S> suppliers. Source extracted <date>.
```

### 5. Confirm live

Wait for the Actions run to go green, then report the URL:
`https://<owner>.github.io/<repo>/`

First-ever deploy: enable Pages in repo Settings → Pages → Source: GitHub Actions. That is a
settings change — ask the user to do it, or to confirm before you do.

---

## What the dashboard shows

Scope is deliberate and has been narrowed over several rounds. Do not widen it without asking:

- **Mapped only.** `fk_product_map IS NOT NULL` at the time of the change. Unmapped rows are out.
- **Real edits only.** `previous_description IS NOT NULL`. Rows where the column was populated for
  the first time are out — they are onboarding, not churn.
- **SKUs, not rows**, is the headline. One product edited across five supplier feeds is one SKU.
- **2026 year to date.**

Panels: edited SKUs by supplier, edited SKUs by month, and a download of the full edit detail.
There is deliberately no edit-frequency panel, no token-churn panel, no latest-edit-per-SKU view,
no supplier-account breakdown, and no inline edit feed. Each was tried and cut. Do not reinstate
them unless asked.

## Data notes that matter

- **Supplier** is `supplier_key` up to the first `-` or `_`. The suffix is usually a state or branch
  (`tempetyres-wa`, `yokohama-qld-brisbane`, `yhi_adelaide`), but it also catches non-location
  variants (`michelin-promo`, `newbee-virtual`, `chapelcorner-mps`), which fold into the parent.
- **Neither `supplier_key` nor `fk_supplier` is unique.** `tyredepot` appears under `fk_supplier` 3
  and 17; `fk_supplier` 17 carries both `tyredepot` and `tempetyres-nsw`. Group on the pair.
- **One upstream edit, several rows.** Each state account holds its own `id_supplier_stock` for the
  same product. Count distinct `spinach_pid` for SKUs, row count for stock records.
- **The current month is partial.** `build.py` compares the data's max `changed_at` against the last
  calendar day of that month. If the data stops short, the month is hatched, excluded from the
  average, and labelled ("1 to 9 September only") from the actual first and last day present. If the
  data reaches the end of the month, nothing is hatched and every month counts toward the average.
  The average line is suppressed entirely below two whole months.
- **The source table does not refresh.** It was backfilled once on 2026-09-10 from the popeye CDC
  stream with no transfer config. Re-running the extract against it returns the same data until
  someone rebuilds the table. If the user expects new data, check the table's freshness first and
  tell them if it has not moved.
- **`fk_product_map` is point-in-time**, the mapping as it stood when the change was captured, not
  today. The dashboard does not claim otherwise and neither should any summary of it.

## Do not

- Push without asking, or treat an earlier approval as covering a later run.
- Change repo settings, visibility, or Pages config without the user's say-so.
- Build on an extract whose counts did not reconcile.
- Hand-edit `docs/index.html`. It is generated — change `templates/dashboard.html` or
  `scripts/build.py` and rebuild, or the next run silently reverts the fix.
- Hardcode a count, a date or a month name in `templates/dashboard.html`. Every figure the page
  shows comes from the `__DATA__` payload (`stats` and `meta`), so a rebuild updates all of it. Text
  typed into the template survives the rebuild and will quietly contradict the charts.
- Commit BigQuery credentials or a service-account key. The extract runs on the caller's own
  `gcloud`/`bq` auth.
