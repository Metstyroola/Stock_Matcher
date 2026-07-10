Real snapshot data for the Stock Matching dashboard (betty-boop).

These two files are the REAL aggregate snapshot referenced by
apps/dashboards/web/dashboards/stock-matching/MANIFEST.md (hashes pinned there).
They are NOT committed to git (the platform gitignores web/dashboards/*/data/).

To make the dashboard show real data on the site, stage them to GCS so the
container build rsyncs them into data/ (needs bucket access — @alepuri or an
analyst with access):

  gsutil cp meta.js suppliers.js gs://betty-boop-data/stock-matching/

MANIFEST hashes (must match exactly):
  meta.js       466 bytes    sha256 8ec324019863094bd6d6caf5039d4a40ddf897707bff8bc67203dd6447290ed5
  suppliers.js  18229 bytes  sha256 dd9104b42c53297454538a9e72fc9fd3fd4654c8503e925c04e4adf05775b205
