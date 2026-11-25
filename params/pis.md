# PIs parameters

Rows of the sheet `PIs`: the dates of each Program Increment, used by the data check DC18 (the sprints of the stories of
a feature must fall inside the PI of the feature). **Empty in the repo**: the calendar is the client's. Type the rows in
the workbook, with real dates.

`PI` is the value of the PI field (`PI#XX-YYYY`), one PI per row. `Start` is included, `End` is excluded. Dates are UTC dates
(midnight), for example 2026-10-05 and 2027-02-08 (a PI of 18 weeks). A PI with no `Start` and no `End` is left out of the
metrics, as a PI the sheet lacks (the status says `PIs without dates`); one date without the other is refused.

A PI is the long cycle of the Program: it holds several iterations (sheet `Iterations`), for example 6 of 3 weeks.

| PI | Start | End |
|---|---|---|
