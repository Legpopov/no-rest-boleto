# Price history

`python -m collector.history --git-ref origin/main` imports every available
`data/prices.json` revision. `python -m collector.history data/prices.json`
archives a new snapshot; the collection workflow runs it after merge.

The archive is partitioned by **observation month** in `data/history/*.jsonl.gz`.
Each JSON line contains its source, currency, actual observation timestamp,
origin, destination, and original successful source payload. Hotel observations
are independent of origin. The SHA-256 identity makes repeated imports idempotent.
Stale/error blocks and blocks without observation dates do not become new observations.
The deterministic gzip files and `index.json` can be read with Python's standard library.

Collection month is not travel month. Flight `by_month` contains departure-month
statistics, while legacy `by_week` is kept as a weekly statistic. Neither is a
complete transaction history. Undated rail/hotel prices remain undated. A current
quote for a future departure is an observed quote, not a completed sale.

Replaying this repository cannot recover periods before its first snapshot.
Imported observations must remain separate from reconstructed prices and forecasts.
No transport preference, company filter, allowance or workbook belongs in this archive.
