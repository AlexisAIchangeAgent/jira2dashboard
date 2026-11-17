"""The report "expected / obtained" of the scenario art-metrics: the numbers worked out BY HAND (tests/art_oracle.py) next to the numbers the
workbook shows, one line each, OK or DIFF. It goes into the page index.html of the charts, so that one page tells what to look at and
what is wrong. Written by the Excel test (`art_metrics_excel.feature`) and by `mock-workbook --scenario art-metrics --run`.
Pure but for `report_for_workbook`, which reads the workbook through COM. The oracle stays in tests/: the delivered code never imports it.
"""
import html

from art_oracle import (ART_ROWS, DONE_A, DONE_B, FLOW_EFFICIENCY, FLOW_TIME_DAYS, PIS, PLANNED, TEAM_ROWS, VELOCITY)
from jira2dashboard.flow.model import NOT_COMPUTABLE

ART_COLUMNS = ("Engaged", "Added", "Deferred", "Released", "ReleasedAdded", "CarryOver")
TEAM_COLUMNS = ("Engaged", "Added", "Deferred", "Released")
TOLERANCE = 1e-9


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
    """The checks, as [(title, [row])]. `tables` = the rows (dicts by column name) of PIMetrics, IterationFeatures and Metrics."""
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
    return [("Feature Predictability of each PI", predictability),
            (f"Features of the ART per iteration ({' / '.join(ART_COLUMNS)})", per_iteration),
            (f"Features of each team, PI-3 and PI-4 ({' / '.join(TEAM_COLUMNS)})", per_team),
            ("Velocity of the ART per iteration (Features released)", velocity),
            ("Flow Time median / p85 (days) and Flow Efficiency of the ART per iteration", flow)]


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
    for name in ("PIMetrics", "IterationFeatures", "Metrics"):
        header, body = sheets.read_table(wb, name, name)
        tables[name] = [dict(zip(header, r)) for r in body]
    return render(sections(tables, predictability_mode(wb)))
