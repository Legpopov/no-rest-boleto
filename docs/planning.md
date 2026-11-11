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

The Excel scenario repeats the individual route/month 2025–2026 reference rate
in each year after 2026. This explicit scenario is distinct from the primitive
API, which accepts an independently sourced index for each year. The catalog
keeps matched, unambiguous one-way RUB air references with source links;
surface reserve estimates are explicitly labelled separately.
There is currently no calibrated peer/official transport-index fallback for air.
Rail/bus have a distinct reserve method: mean undated/timetable reference prices,
flat seasonality, and a 4% general-inflation scenario sourced to the Bank of Russia
forecast of 5 November 2026. This is not historical transport inflation. Beyond
2029 it extrapolates the scenario. Quantitative forecast error is unknown.
Each rail class/firm group retains its sample count; Excel applies the workbook's
existing class rules and weights groups by their counts. No organization-specific
class restrictions are embedded in the published data. A bus is checked
independently of whether a train is available. Source errors, missing price,
missing page and source-confirmed absence of a direct bus have distinct statuses.
Unknown endpoints return an explicit outside-coverage message and `#N/A`, never
an automatically substituted Moscow or neighboring-city fare.
Unknown semester months propagate `#N/A` into dependent budget sums rather than
being silently treated as zero or averaged away. Known zero hotel nights are a
valid zero expense; missing prices are not.
