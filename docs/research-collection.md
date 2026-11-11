# Historical transport research

Use the sequence: inspect 2–3 real examples, download raw sources, report coverage,
parse by headers, then verify extracted rows and repeat the import. Keep failures
distinct from confirmed absence. For changing URLs inspect dated archive versions;
do not infer historic prices from a current search cache.

The expanded snapshot inspects 24 original ATOR publications. Raw HTML and SHA-256
receipts are retained locally, outside the repository. Reproduce with:

```
python -m collector.research /path/to/manifest.json --out data/research
```

The manifest records id, URL, retrieval date, HTTP status and SHA-256. Parsing is
offline and fails on an unexpected currency, origin or destination. Empty fares,
unresolved route captions and explicitly unsupported fare columns are recorded
in `rejections.json`; a summary table is not a route observation.
`sources.json` reports per-source coverage; failed downloads are not observations.

`ator-quotes.json` contains one-way minimum fares on reference departure dates,
separately with and without baggage. Airline, date exceptions and source hashes
are retained. These are **not monthly mean transaction prices**. Do not average
them together with round trips, sold-ticket means or unspecified baggage classes.
Comparisons between reference dates are monitoring proxies for month-to-month
movement, and must be labelled as such. A few months are not a full annual series.

Monthly official CPI, monthly average route prices, and dated minimum-fare
monitoring are separate evidence types. Official CPI may be used as the final
fallback, never presented as observed inflation for a particular route. Regional
prices per 100/1000 km must not be labelled as actual endpoint-to-endpoint fares.

`ator-route-yoy.json` compares matching origin, destination, baggage, currency,
one-way basis and calendar month across years. These remain **reference-date**
comparisons; check booking horizons in the linked publications before adopting a
rate for a forecast. `comparison-exclusions.json` retains the IDs of observations
excluded because more than one reference quote exists for a route/month. This
occurs when a publication carries a date exception referring to an earlier month.
The underlying observations are preserved rather than overwritten or relabelled.

It also occurs for January 2026: ATOR measured both 1 and 12 January. The source
has different holiday/demand and booking-horizon conditions. Both dates are
retained; neither is silently selected to represent the entire month.

As of the snapshot retrieved on 10 November 2026 UTC (11 November Moscow),
there are 2,211 reference fares and 881 matched 2025→2026 comparisons, including
baggage variants. The matched months are March and May–December. These are
advance-fare observations: December travel dates were already quoted in November.
`coverage.json` lists all 24 months, including months with zero extracted
observations, and comparison exclusions. Zero extraction is a statement about
this snapshot, not proof that a publication or price never existed.

Newer November/December articles put the common departure date outside the city
captions. Two source-ID-specific mappings retain those dates and the asterisk
exceptions (3 November / 2 December). The November double asterisk denotes a
fare already including baggage, not a departure-date change: its purported
no-baggage column is excluded. In August 2025, Kaliningrad's caption omits the day;
the date is explicitly taken from the section heading (1 August). Two cells omit
the RUB abbreviation; their unit is taken from surrounding cells of the same
inspected table. These repairs carry `date_method` / `currency_method` flags.

Discovery used the publisher's linked monthly studies and tickets archive
(`/taxonomy/term/430`, pages 0–6), not only search results. The public sitemap
contains mostly numeric node URLs and did not establish topic coverage by itself.
The July summary and the July table publication are distinct sources.

The source matrix must cover 2025 and available 2026 months for each mode. Search
publisher report archives, official statistical datasets and historical snapshots;
record access failures and missing periods. Do not force a numeric forecast by
inventing an annual coefficient when only a month-to-month comparison exists.
