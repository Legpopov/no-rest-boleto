# Planning calculation contract

`collector.planning` provides calculation primitives; it is not yet connected to
the workbook or the source archive. Source calibration and workbook integration
must supply the inputs below before an end-to-end estimate is available.

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
