# PIs parameters

Rows of the sheet `PIs`: the dates of each Program Increment, used by the data check DC18 (the sprints of the stories of
a feature must fall inside the PI of the feature). **Empty in the repo**: the calendar is the client's. Type the rows in
the workbook, with real dates.

`PI` is the value of the PI field, any name (for example `PI#01-2026` or `PI#4.1-2024`), one PI per row; the PIs come by `Start`,
those without dates after them in the natural order of their names (PI#3-2024 before PI#10-2026). `Start` is included, `End` is excluded. Dates are UTC dates
(midnight), for example 2026-10-05 and 2027-02-08 (a PI of 18 weeks). A PI with no `Start` and no `End` is left out of the
metrics, as a PI the sheet lacks (the status says `PIs without dates`); one date without the other is refused.

A PI is the long cycle of the Program: it holds several iterations (sheet `Iterations`), for example 6 of 3 weeks.

| PI | Start | End |
|---|---|---|
