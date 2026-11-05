---
name: steering-explain
description: Explain the ART's flow metrics from the Confluence page "Flow metrics <PI>" published by the RTE. Gives values, trends, figures that are not computable, and the data checks behind them. Never computes a metric.
---

# Explain the flow metrics

1. Open the page "Flow metrics <PI>" the person names, or ask which one. Read its header: as-of date,
   FlowLoadMode, scope.
2. Quote values from tables A and B exactly as written, with the period, the team and the as-of date. Never
   recompute, round or average them.
3. Trend: compare a value only with the same team's previous period on the page. Say up, down or same, and quote
   both values. If the previous period is not on the page, give no trend.
4. For a "not computable" value, give the reason from the table below. Add the checks of table C that affect that
   metric and have the Result KO.
5. List the checks whose Result is KO, with Items, Severity and the metric they affect. A Result of "not
   computable" means a field id is missing in the workbook's Config.
6. If a team, a period or a metric is missing from the page, write "not on the page".

| Metric | Not computable when | Checks |
|---|---|---|
| Velocity | never (0 is a value) | DC06, DC07, DC08 |
| FlowTimeMedian, FlowTimeP85 | no item done in the period has a valid start | DC09, DC10 (items left out: FlowTimeExcluded), DC06, DC07 |
| FlowLoadAverage, FlowLoadEnd | the period starts after the as-of date | DC08, DC06 |
| FlowEfficiency | as Flow Time, or time spent in a status missing from Workflows | DC08 |
| DistDone<bucket> | Velocity is 0 | an empty work type counts as Unmapped; DC25 |
| FlowPredictability | always: Jira holds no PI objectives with planned and actual business value | DC21 |

Output: one short paragraph per metric asked about. Last line: "Source: page <title>, as of <date>".
