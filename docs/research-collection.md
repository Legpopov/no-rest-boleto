# Historical transport research

Use the sequence: inspect 2–3 real examples, download raw sources, report coverage,
parse by headers, then verify extracted rows and repeat the import. Keep failures
distinct from confirmed absence. For changing URLs inspect dated archive versions;
do not infer historic prices from a current search cache.

The first pilot imports three original ATOR publications. Raw HTML and SHA-256
receipts are retained locally, outside the repository. Reproduce with:

```
python -m collector.research /path/to/manifest.json --out data/research
```

The manifest records id, URL, retrieval date, HTTP status and SHA-256. Parsing is
offline and fails on an unexpected header, currency, origin or destination.
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

The source matrix must cover 2025 and available 2026 months for each mode. Search
publisher report archives, official statistical datasets and historical snapshots;
record access failures and missing periods. Do not force a numeric forecast by
inventing an annual coefficient when only a month-to-month comparison exists.
