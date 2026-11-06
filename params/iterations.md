# Iterations parameters

Rows of the sheet `Iterations`: the short cycles over which the flow metrics are computed. An iteration sits inside
a PI (the long cycle, sheet `PIs`): a PI of 18 weeks holds 6 iterations of 3 weeks, a PI of 12 weeks holds 6
iterations of 2 weeks. **Empty in the repo**: the calendar is the client's. Type the rows in the workbook, with
real dates.

An iteration is half-open: `Start` is included, `End` is excluded (an item done on `End` belongs to the next
iteration). Dates are UTC; `Start` and `End` are dates (midnight), for example 2026-10-05 and 2026-10-26.

| Iteration-label | Start | End |
|---|---|---|
