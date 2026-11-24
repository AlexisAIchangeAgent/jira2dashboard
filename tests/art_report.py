"""The report "expected / obtained" of the scenario art-metrics: the numbers worked out BY HAND (tests/art_oracle.py) next to the numbers the
workbook shows, one line each, OK or DIFF. It goes into the page index.html of the charts, so that one page tells what to look at and
what is wrong. Written by the Excel test (`art_metrics_excel.feature`) and by `mock-workbook --scenario art-metrics --run`.
Pure but for `report_for_workbook`, which reads the workbook through COM. The oracle stays in tests/: the delivered code never imports it.
"""
import html

from art_oracle import (ART_ROWS, CHAIN, DONE_A, DONE_B, FLOW_DISTRIBUTION_CHECK, FLOW_EFFICIENCY, FLOW_EFFICIENCY_CHECK, FLOW_LOAD_CHECK,
                        FLOW_TIME_CHECK, FLOW_TIME_DAYS, PI_ITEMS, PIS, PLANNED, STATUS_AVERAGE_DAYS, STATUS_SHARE, TEAM_ROWS, VELOCITY)
from jira2dashboard.flow.metrics import BUCKETS
from jira2dashboard.flow.model import NOT_COMPUTABLE

ART_COLUMNS = ("Engaged", "Added", "Deferred", "Released", "ReleasedAdded", "CarryOver")
TEAM_COLUMNS = ("Engaged", "Added", "Deferred", "Released")
TOLERANCE = 1e-9


def _flat(values: tuple) -> tuple:
    """(Done, (counts...), Gap, Reason, Note) of FLOW_DISTRIBUTION_CHECK as one row of figures."""
    done, counts, *rest = values
    return (done, *counts, *rest)


# The accounting tables of the flow metrics (AL, 2026-11-22), the ART per iteration: (table, what it tells, columns, expected by label, the
# position of a percentage). The oracle is the one the Excel scenarios assert, typed by hand in tests/art_oracle.py.
CHECK_TABLES = (
    ("FlowTimeCheck", "why the Flow Time is what it is", ("Done", "ExclReopened", "ExclNoStartRule", "ExclNeverStarted", "ExclBackToBacklog", "Counted", "Gap", "Reason"),
     FLOW_TIME_CHECK, None),
    ("FlowEfficiencyCheck", "why the Flow Efficiency is what it is",
     ("Counted", "WithUnclassified", "ActiveDays", "WaitDays", "WaitStatuses", "Efficiency", "Gap", "Reason"), FLOW_EFFICIENCY_CHECK, 5),
    ("FlowLoadCheck", "why the Flow Load is what it is", ("Days", "DaysMeasured", "StartStatuses", "Gap", "Reason"), FLOW_LOAD_CHECK, None),
    ("FlowDistCheck", "what the Flow Distribution counts", ("Done", *BUCKETS, "Gap", "Reason", "Note"),
     {label: _flat(values) for label, values in FLOW_DISTRIBUTION_CHECK.items()}, None),
)


def same(expected, obtained) -> bool:
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(obtained, (int, float)) and abs(obtained - expected) < TOLERANCE
    return expected == obtained


def show(value) -> str:
    if isinstance(value, float) and not value.is_integer():
        return f"{value:.4g}"
    if isinstance(value, (int, float)):
        return str(int(value))
    return "-" if value in (None, "") else str(value)


def percent(value) -> str:
    return f"{value * 100:.2f} %" if isinstance(value, (int, float)) else show(value)


def row(label: str, expected: tuple, obtained: tuple, formats: tuple = ()) -> tuple[str, str, str, bool]:
    """(label, expected text, obtained text, ok): the two tuples are compared one figure after the other; `formats` says how to
    show the figure at each position (default `show`)."""
    ok = len(expected) == len(obtained) and all(same(e, o) for e, o in zip(expected, obtained))

    def text(values):
        return " / ".join((formats[i] if i < len(formats) else show)(v) for i, v in enumerate(values))
    return label, text(expected), text(obtained), ok


def by_key(rows: list[dict], *keys: str) -> dict:
    return {tuple(r[k] for k in keys): r for r in rows}


def sections(tables: dict[str, list[dict]], mode: str = "A") -> list[tuple[str, list[tuple]]]:
    """The checks, as [(title, [row])]. `tables` = the rows (dicts by column name) of PIMetrics, IterationFeatures, Metrics, PIStatusTime and
    the accounting tables of CHECK_TABLES."""
    pi_rows = by_key(tables["PIMetrics"], "PI")
    done = DONE_A if mode == "A" else DONE_B
    predictability = []
    for pi in PIS:
        got = pi_rows.get((pi,), {})
        predictability.append(row(f"{pi}: planned / done (mode {mode}) / predictability",
                                  (PLANNED[pi], done[pi], done[pi] / PLANNED[pi]),
                                  (got.get("FeaturesPlanned"), got.get("FeaturesDone"), got.get("FeaturePredictability")),
                                  (show, show, percent)))
    art = by_key([r for r in tables["IterationFeatures"] if r["Team"] == "ART"], "Iteration-label")
    per_iteration = [row(label, expected, tuple(art.get((label,), {}).get(c) for c in ART_COLUMNS)) for label, expected in ART_ROWS.items()]
    teams = by_key(tables["IterationFeatures"], "Team", "Iteration-label")
    per_team = [row(f"{team} {label}", expected, tuple(teams.get((team, label), {}).get(c) for c in TEAM_COLUMNS))
                for team, rows in TEAM_ROWS.items() for label, expected in rows.items()]
    metrics = by_key([r for r in tables["Metrics"] if r["Team"] == "ART"], "Iteration-label")
    velocity = [row(label, (VELOCITY[label],), (metrics.get((label,), {}).get("Velocity"),)) for label in ART_ROWS]
    flow = []
    for label in ART_ROWS:
        got = metrics.get((label,), {})
        if VELOCITY[label]:
            efficiency = FLOW_EFFICIENCY["PI-" + label[3]]
            flow.append(row(label, (FLOW_TIME_DAYS, FLOW_TIME_DAYS, efficiency),
                            (got.get("FlowTimeMedian"), got.get("FlowTimeP85"), got.get("FlowEfficiency")),
                            (show, show, percent)))
        else:
            flow.append(row(label, (NOT_COMPUTABLE, NOT_COMPUTABLE), (got.get("FlowTimeMedian"), got.get("FlowEfficiency"))))
    status_time = by_key(tables["PIStatusTime"], "PI", "Status")
    cycle = [row(f"{pi} {status}", (PI_ITEMS[pi], STATUS_AVERAGE_DAYS[pi][status], STATUS_SHARE[pi][status]),
                 tuple(status_time.get((pi, status), {}).get(c) for c in ("Items", "AverageDays", "Share")), (show, show, percent))
             for pi in PIS for status in CHAIN]
    return [("Feature Predictability of each PI", predictability),
            (f"Features of the ART per iteration ({' / '.join(ART_COLUMNS)})", per_iteration),
            (f"Features of each team, PI-3 and PI-4 ({' / '.join(TEAM_COLUMNS)})", per_team),
            ("Velocity of the ART per iteration (Features released)", velocity),
            ("Flow Time median / p85 (days) and Flow Efficiency of the ART per iteration", flow),
            ("Cycle time by status, per PI (Features released / average days per Feature / share of the cycle time)", cycle),
            *(check_section(tables, *table) for table in CHECK_TABLES)]


def check_section(tables: dict, table: str, tells: str, columns: tuple, expected: dict, percent_at: int | None) -> tuple[str, list[tuple]]:
    """One accounting table of the ART per iteration; an empty text read back from a file is None: shown and compared as ""."""
    got = by_key([r for r in tables.get(table, []) if r["Team"] == "ART"], "Iteration-label")
    formats = tuple(percent if i == percent_at else show for i in range(len(columns)))
    rows = [row(label, values, tuple("" if (v := got.get((label,), {}).get(c)) is None else v for c in columns), formats)
            for label, values in expected.items()]
    return f"{table} of the ART per iteration: {tells} ({' / '.join(columns)})", rows


def render(found: list[tuple[str, list[tuple]]]) -> str:
    """The HTML fragment: a summary line, then a table per section, the lines that differ marked DIFF and shown first."""
    e = html.escape
    total = sum(len(rows) for _title, rows in found)
    wrong = sum(1 for _title, rows in found for r in rows if not r[3])
    verdict = (f'<p class="verdict ok">All {total} figures are the ones worked out by hand.</p>' if not wrong else
               f'<p class="verdict diff">{wrong} of {total} figures differ from the ones worked out by hand.</p>')
    out = [verdict]
    for title, rows in found:
        bad = sum(1 for r in rows if not r[3])
        out.append(f'<details class="rep"{" open" if bad else ""}><summary>{e(title)} - {len(rows) - bad} of {len(rows)} OK</summary>'
                   '<table><tr><th>Check</th><th>Expected (by hand)</th><th>Obtained (workbook)</th><th></th></tr>')
        for label, expected, obtained, ok in sorted(rows, key=lambda r: r[3]):          # the lines that differ first
            out.append(f'<tr class="{"ok" if ok else "diff"}"><td>{e(label)}</td><td>{e(expected)}</td><td>{e(obtained)}</td>'
                       f'<td>{"OK" if ok else "DIFF"}</td></tr>')
        out.append("</table></details>")
    return "".join(out)


def predictability_mode(wb) -> str:
    """The value of the Config parameter PredictabilityMode of the open workbook (A when it cannot be read)."""
    from jira2dashboard import sheet_names as S
    try:
        for name, value, *_rest in wb.Worksheets(S.of("Config")).UsedRange.Value:
            if name == "PredictabilityMode":
                return str(value or "A").strip().upper()
    except Exception:
        pass
    return "A"


def report_for_workbook(wb) -> str:
    """Read the tables of the open workbook and return the HTML report (COM)."""
    from jira2dashboard.flow import sheets
    tables = {}
    for name in ("PIMetrics", "IterationFeatures", "Metrics", "PIStatusTime", *(t[0] for t in CHECK_TABLES)):
        header, body = sheets.read_table(wb, name, name)
        tables[name] = [dict(zip(header, r)) for r in body]
    return render(sections(tables, predictability_mode(wb)))
