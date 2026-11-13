# Planning calculation contract

`collector.planning` provides calculation primitives.
`collector.planning_catalog` publishes matched source fares to `data/planning.json`;
`excel/modPlanning.bas` connects these to the existing trip sheet. Workbook
installation and present coverage limitations are documented in `excel/README.md`.

The VBA integration has three boundaries: `modPlanning` adapts ranges and worksheet
errors, `modPlanningCore` evaluates ordinary arrays/scalars without Excel I/O, and
`modPlanningStorage` validates the catalog and handles safe writes/backups. Column
positions have shared names in `PlanningColumn`. Existing worksheet function
signatures remain compatible. Both Python and native Excel execute the cases in
`tests/fixtures/planning-contract.json`; the implementations remain independent
because recipients need an offline workbook without a Python runtime.

- `semester_months(year, cohort, semester, Calendar(...))` resolves calendar months.
  The editable calendar defaults to September–January and February–June.
- `annual_index(current, previous)` calculates a decimal annual change. The caller
  must first establish comparable month, class, baggage, currency and booking horizon.
- `select_index(route, peers, official)` chooses sourced evidence in that order.
  Peers must already be matched by the caller; the function does not infer similarity.
  The median of peer rates gives each peer route equal weight.
- `forecast_month` compounds explicit annual indices. Reusing an historical index
  for multiple future years is a forecast assumption that must be explicit in each
  supplied year's record, with its source. Missing rates do not silently mean 0%.
- `forecast_semester` uses equal month weights. An incomplete semester returns
  `insufficient_data` and `price: null`, with coverage and month-level reasons,
  rather than an understated average of available months.
- `choose_transport` follows the user's preference order, skipping unavailable
  modes. It does not replace preference with a cheapest-price decision.
- `route_total` requires connected, explicitly one-way segments. It never divides
  a round-trip airfare by two or doubles a chain to invent a return journey.

Inputs are market prices and sourced indices. No internal workbook allowances or
organization-specific filters are embedded in these functions. Hotel nights,
per-person quantities, transit stays and expense-category aggregation belong to
the integration layer, and must not be inferred from transport prices.

Run checks with `python -m unittest discover -s tests -v`.

The Excel catalog is schema 3. `data/planning.json` keeps inspectable records;
`data/planning-excel.json` losslessly encodes the same worksheet fields with a string
dictionary and fixed columns. `PlanningCatalogMatrix` validates decoded types before
any write. Python verifies equality for every published record.

Current cached one-way air references take precedence over matched ATOR references
for the same route/month when their travel year does not exceed the target year.
Unknown baggage requires the workbook's explicit one-way allowance; without it a
known-baggage historical reference or an unavailable estimate is used. A future
departure-year reference is never backdated as an earlier observation. Through fares
may already include connections; separately entered legs must not double count them.

For geographically eligible domestic routes the national all-passenger-transport
YoY index is a broad proxy repeated as an explicit scenario, not a mode/route index
or an official forecast. Foreign/excluded routes retain the sourced RUB macro
scenario, with no separate FX/local-price model. Historical ATOR route/month rates
remain explicit repeated scenarios. Surface seasonality is still flat.

Hotel bases are calculated from fresh source items with the workbook's existing
filters, not from potentially retained BR cells. Their observation year/month,
price/index sources and method are saved offline. Hotel forecasts compound the
annual scenario by elapsed calendar months and average over the semester, then
multiply by people and explicit nights. Russian hotel indices are not applied to
foreign or excluded destinations. Transit without nights remains zero.

The official branch-city selector defines 50 origins including Moscow. Destinations
also include source-returned city codes resolved against the provider's city
dictionary; no nearest airport is silently substituted. Geographic scope is distinct
from fare and semester coverage. Source failures and an unconfirmed station do not
prove that a route does not operate. Coverage is recorded per directed route/mode;
surface snapshot completeness is checked against its explicit requested-route set,
not a fixed number of cities. Missing prices/months propagate as `#N/A` into totals.
