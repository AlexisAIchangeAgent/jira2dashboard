# Periods parameters

Rows of the sheet `Periods`: the iterations (or PIs) over which the flow metrics are computed. **Empty in the
repo**: the calendar is the client's. Type the rows in the workbook, with real dates.

A period is half-open: `Start` is included, `End` is excluded (an item done on `End` belongs to the next period).
Dates are UTC; `Start` and `End` are dates (midnight), for example 2026-10-05 and 2026-10-19.

| Period | Start | End |
|---|---|---|
