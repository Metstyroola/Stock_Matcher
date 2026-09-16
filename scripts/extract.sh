#!/usr/bin/env bash
# Extract mapped description edits from BigQuery into data/raw_edits.json.
#
# The query tool caps at 3,000 rows per call and truncates silently, so this pulls
# month-sized chunks, splits any month that comes back full, and reconciles the total
# against an independent COUNT(*) before writing anything.
set -euo pipefail

PROJECT="${BQ_PROJECT:-heroic-ruler-198603}"
TABLE="${BQ_TABLE:-heroic-ruler-198603.popeye_mirror.supplier_description_change_log}"
YEAR_START="${YEAR_START:-2026-01-01}"
CAP=3000

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT="$ROOT/data/raw_edits.json"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
mkdir -p "$ROOT/data"

COLS='store_country, id_supplier_stock, fk_supplier, supplier_key, supplier_pid,
      fk_product_map, spinach_pid, previous_description, new_description,
      FORMAT_TIMESTAMP("%F %H:%M:%S", previous_seen_at) AS previous_seen_at,
      FORMAT_TIMESTAMP("%F %H:%M:%S", changed_at) AS changed_at'

# Mapped, genuine edits only. Keep in step with publish-dashboard.md ("What the dashboard shows").
WHERE='fk_product_map IS NOT NULL AND previous_description IS NOT NULL'

bq_json() {  # $1 = SQL
  bq query --project_id="$PROJECT" --use_legacy_sql=false --format=prettyjson --max_rows=100000 "$1"
}

echo "Reconciliation target..."
EXPECTED=$(bq_json "SELECT COUNT(*) AS n FROM \`$TABLE\`
  WHERE changed_at >= TIMESTAMP('$YEAR_START') AND $WHERE" | python3 -c 'import json,sys; print(json.load(sys.stdin)[0]["n"])')
echo "  expecting $EXPECTED rows"

MONTHS=$(bq_json "SELECT DISTINCT FORMAT_TIMESTAMP('%Y-%m', changed_at) AS m FROM \`$TABLE\`
  WHERE changed_at >= TIMESTAMP('$YEAR_START') AND $WHERE ORDER BY m" \
  | python3 -c 'import json,sys; print(" ".join(r["m"] for r in json.load(sys.stdin)))')

i=0
for M in $MONTHS; do
  START="${M}-01"
  END=$(python3 -c "
import datetime as d
y,m=map(int,'$M'.split('-'))
print(d.date(y+(m==12), 1 if m==12 else m+1, 1))")
  echo "  $M ..."
  bq_json "SELECT $COLS FROM \`$TABLE\`
    WHERE changed_at >= TIMESTAMP('$START') AND changed_at < TIMESTAMP('$END') AND $WHERE
    ORDER BY changed_at" > "$TMP/chunk_$i.json"

  N=$(python3 -c "import json;print(len(json.load(open('$TMP/chunk_$i.json'))))")
  if [ "$N" -ge "$CAP" ]; then
    # Month hit the cap — it is truncated. Re-pull it in day-sized slices.
    echo "    $M returned $N rows (at the cap) — re-pulling by day"
    rm "$TMP/chunk_$i.json"
    D="$START"
    while [ "$D" \< "$END" ]; do
      NEXT=$(python3 -c "
import datetime as d
print(d.date.fromisoformat('$D') + d.timedelta(days=1))")
      bq_json "SELECT $COLS FROM \`$TABLE\`
        WHERE changed_at >= TIMESTAMP('$D') AND changed_at < TIMESTAMP('$NEXT') AND $WHERE
        ORDER BY changed_at" > "$TMP/chunk_$i.json"
      DN=$(python3 -c "import json;print(len(json.load(open('$TMP/chunk_$i.json'))))")
      [ "$DN" -ge "$CAP" ] && { echo "ERROR: $D alone exceeds the $CAP cap — split by hour"; exit 1; }
      i=$((i+1)); D="$NEXT"
    done
    continue
  fi
  i=$((i+1))
done

python3 - "$TMP" "$OUT" "$EXPECTED" <<'PY'
import json, sys, glob, os, datetime
tmp, out, expected = sys.argv[1], sys.argv[2], int(sys.argv[3])
rows, cols = [], None
for f in sorted(glob.glob(os.path.join(tmp, 'chunk_*.json'))):
    data = json.load(open(f))
    if not data:
        continue
    if cols is None:
        cols = list(data[0].keys())
    rows += [[r.get(c) for c in cols] for r in data]
seen, dedup = set(), []
for r in rows:
    k = tuple(r)
    if k not in seen:
        seen.add(k); dedup.append(r)
dedup.sort(key=lambda r: r[cols.index('changed_at')], reverse=True)
if len(dedup) != expected:
    print(f'RECONCILIATION FAILED: assembled {len(dedup):,}, expected {expected:,}')
    print('Do not build on this extract. Narrow the chunks and re-run.')
    sys.exit(1)
# Stamp the extract date here, not at build time: mtime is the checkout date on a
# fresh clone, and the page prints this in its footer.
extracted = datetime.datetime.now(datetime.timezone.utc).date().isoformat()
json.dump({'cols': cols, 'rows': dedup, 'extracted': extracted},
          open(out, 'w'), separators=(',', ':'))
print(f'OK: {len(dedup):,} rows reconciled -> {out}')
PY
