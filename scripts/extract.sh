#!/usr/bin/env bash
# Extract mapped description edits from BigQuery into data/raw_edits.json.
#
# Every query path has a row limit and some truncate silently, so this pulls month-sized
# chunks, splits any month that comes back at the limit, and reconciles the total against
# an independent COUNT(*) before writing anything.
#
# MAX_ROWS is the limit in force. It is bq's --max_rows here; if you run the same SQL
# through the BigQuery MCP connector instead, that caps at 3,000 per call, so set
# MAX_ROWS=3000 to keep the truncation check honest.
set -euo pipefail

PROJECT="${BQ_PROJECT:-heroic-ruler-198603}"
TABLE="${BQ_TABLE:-heroic-ruler-198603.popeye_mirror.supplier_description_change_log}"
YEAR_START="${YEAR_START:-2026-01-01}"
MAX_ROWS="${MAX_ROWS:-100000}"
CAP="$MAX_ROWS"

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

bq_json() {  # $1 = SQL -> JSON on stdout
  bq query --project_id="$PROJECT" --use_legacy_sql=false --format=prettyjson \
    --max_rows="$MAX_ROWS" "$1"
}

# bq prints nothing at all for a zero-row result, which json.load cannot parse.
# A day with no edits is entirely normal, so normalise that to an empty array.
bq_json_to() {  # $1 = destination file, $2 = SQL
  bq_json "$2" > "$1"
  grep -q '[^[:space:]]' "$1" 2>/dev/null || echo '[]' > "$1"
}

# Zero-pad so the chunk files sort in pull order rather than chunk_10 before chunk_2.
chunk() { printf '%s/chunk_%04d.json' "$TMP" "$1"; }

echo "Reconciliation target..."
EXPECTED=$(bq_json "SELECT COUNT(*) AS n FROM \`$TABLE\`
  WHERE changed_at >= TIMESTAMP('$YEAR_START') AND $WHERE" | python3 -c 'import json,sys; print(json.load(sys.stdin)[0]["n"])')
echo "  expecting $EXPECTED rows"

MONTHS=$(bq_json "SELECT DISTINCT FORMAT_TIMESTAMP('%Y-%m', changed_at) AS m FROM \`$TABLE\`
  WHERE changed_at >= TIMESTAMP('$YEAR_START') AND $WHERE ORDER BY m" \
  | python3 -c 'import json,sys; d=sys.stdin.read().strip(); print(" ".join(r["m"] for r in (json.loads(d) if d else [])))')

if [ -z "${MONTHS// /}" ]; then
  echo "No rows match the scope from $YEAR_START onwards — nothing to extract." >&2
  exit 1
fi

i=0
for M in $MONTHS; do
  START="${M}-01"
  END=$(python3 -c "
import datetime as d
y,m=map(int,'$M'.split('-'))
print(d.date(y+(m==12), 1 if m==12 else m+1, 1))")
  echo "  $M ..."
  F="$(chunk $i)"
  bq_json_to "$F" "SELECT $COLS FROM \`$TABLE\`
    WHERE changed_at >= TIMESTAMP('$START') AND changed_at < TIMESTAMP('$END') AND $WHERE
    ORDER BY changed_at"

  N=$(python3 -c "import json;print(len(json.load(open('$F'))))")
  if [ "$N" -ge "$CAP" ]; then
    # Month came back at the row limit, so it may be truncated. Re-pull it by day.
    echo "    $M returned $N rows (at the $CAP limit) — re-pulling by day"
    rm "$F"
    D="$START"
    while [ "$D" \< "$END" ]; do
      NEXT=$(python3 -c "
import datetime as d
print(d.date.fromisoformat('$D') + d.timedelta(days=1))")
      DF="$(chunk $i)"
      bq_json_to "$DF" "SELECT $COLS FROM \`$TABLE\`
        WHERE changed_at >= TIMESTAMP('$D') AND changed_at < TIMESTAMP('$NEXT') AND $WHERE
        ORDER BY changed_at"
      DN=$(python3 -c "import json;print(len(json.load(open('$DF'))))")
      [ "$DN" -ge "$CAP" ] && { echo "ERROR: $D alone reaches the $CAP limit — split by hour"; exit 1; }
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
# No dedup here. The chunks are half-open, non-overlapping windows, so a row cannot
# arrive twice; dropping duplicates could only ever discard rows that genuinely repeat
# in the source and turn a correct extract into a failed reconciliation. Comparing the
# raw count against COUNT(*) instead catches both a short pull and a double-counted one.
rows.sort(key=lambda r: r[cols.index('changed_at')], reverse=True)
if len(rows) != expected:
    verb = 'short by' if len(rows) < expected else 'over by'
    print(f'RECONCILIATION FAILED: assembled {len(rows):,}, expected {expected:,} '
          f'({verb} {abs(len(rows) - expected):,}).')
    print('Do not build on this extract. A short pull means a chunk was truncated —')
    print('lower MAX_ROWS to force finer slices. An over-count means the windows overlap.')
    sys.exit(1)
# Stamp the extract date here, not at build time: mtime is the checkout date on a
# fresh clone, and the page prints this in its footer.
extracted = datetime.datetime.now(datetime.timezone.utc).date().isoformat()
json.dump({'cols': cols, 'rows': rows, 'extracted': extracted},
          open(out, 'w'), separators=(',', ':'))
print(f'OK: {len(rows):,} rows reconciled -> {out}')
PY
