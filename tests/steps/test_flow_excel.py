"""Steps of features/flow_excel.feature: the Excel formulas against the Python reference, cell by cell."""
import datetime as dt
import gc

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard.excel import com
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import sheets
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import daily_load
from jira2dashboard.flow.metrics import BUCKETS, flow_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Period, Rules
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MockJira
from jira2dashboard.params.model import project_keys
from jira2dashboard.paths import OUT_DIR
from jira2dashboard.testbench import run_vba_tests as rvt

scenarios("../features/flow_excel.feature")

TARGET = OUT_DIR / "flow_workbook.xlsm"
TOLERANCE = 1e-6                    # days (0.09 s): Excel dates are doubles, the reference works in whole seconds


@pytest.fixture(scope="module")
def flow_workbook():
    """A workbook of its own (the shared one is compared with params files by other scenarios)."""
    rvt.build(rvt.ORDER, TARGET)
    holder = {}
    with com.excel_session() as app:
        app.AutomationSecurity = com.MSO_AUTOMATION_SECURITY_LOW
        holder["wb"] = app.Workbooks.Open(str(TARGET.resolve()))
        yield holder
        holder.pop("wb").Close(False)
        del app
        gc.collect()


@pytest.fixture
def world():
    return {}


@pytest.fixture
def jira_server():
    servers = []

    def start(dataset, **options):
        servers.append(MockJira(dataset, **options).start())
        return servers[-1]
    yield start
    for server in servers:
        server.stop()


def set_config(wb, name, value):
    ws = wb.Worksheets("Config")
    row = next(r for r in range(2, 30) if ws.Cells(r, 1).Value == name)
    ws.Cells(row, 2).Value = com.as_cell(value)


def date_or_none(value):
    return None if value in ("", None) else value


def same_moment(excel, expected: dt.datetime | None) -> bool:
    if expected is None:
        return excel in ("", None)
    return excel not in ("", None) and abs(excel - com.to_excel_serial(expected)) < TOLERANCE


def same_number(excel, expected) -> bool:
    if expected is None:
        return excel in ("", None)
    return excel not in ("", None) and not isinstance(excel, str) and abs(excel - expected) < TOLERANCE


def same_figure(excel, expected) -> bool:
    """A metric: the text 'not computable' on both sides, or numbers that agree."""
    if expected == NOT_COMPUTABLE:
        return excel == NOT_COMPUTABLE
    return not isinstance(excel, str) and abs(excel - expected) < TOLERANCE


def table_dicts(wb, name):
    header, rows = sheets.read_table(wb, name, name)
    return [dict(zip(header, row)) for row in rows]


# --- given -------------------------------------------------------------------------------------------

@given(parsers.parse("un faux Jira avec {n:d} tickets générés avec les projets des paramètres"))
def _jira(world, jira_server, n):
    world["dataset"] = generate(seed=11, issues=n, projects=project_keys())
    world["jira"] = jira_server(world["dataset"], credentials=("me@example.com", "tok-123"))


@given(parsers.parse('un classeur dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "{mode}"'))
def _config(world, flow_workbook, mode):
    wb = world["wb"] = flow_workbook["wb"]
    for name, value in {"BaseUrl": world["jira"].base_url, "Email": "me@example.com", "ApiToken": "tok-123",
                        "WorkTypeField": world["dataset"].fields.work_type, "FlowLoadMode": mode,
                        "AsOf": f"{world['dataset'].as_of:%Y-%m-%d %H:%M}"}.items():
        set_config(wb, name, value)


@given(parsers.parse('les itérations de 14 jours depuis le "{first}" jusqu\'à la date de référence dans la feuille Periods'))
def _periods(world, first):
    ws, start, row, periods = world["wb"].Worksheets("Periods"), dt.datetime.fromisoformat(first), 2, []
    while start < world["dataset"].as_of:
        end = start + dt.timedelta(days=14)
        ws.Cells(row, 1).Value = com.as_cell(f"I{row - 1}")
        ws.Cells(row, 2).Value2 = com.to_excel_serial(start)
        ws.Cells(row, 3).Value2 = com.to_excel_serial(end)
        periods.append(Period(f"I{row - 1}", start, end))
        start, row = end, row + 1
    world["periods"] = periods


# --- when --------------------------------------------------------------------------------------------

@when("le VBA extrait les données de flux depuis le faux Jira")
def _extract(world):
    wb = world["wb"]
    wb.Application.Run(f"'{wb.Name}'!modExtract.ExtractFlowData")
    rules, as_of = Rules.from_params(), world["dataset"].as_of
    world["rules"] = rules
    world["items"] = [item_flow(i, as_of, rules) for i in to_flow_issues(world["dataset"])]


@when("les feuilles de flux sont construites avec des formules")
def _build(world):
    sheets.build_flow_sheets(world["wb"])


@when("le tableau de bord est construit")
def _dashboard(world):
    sheets.build_dashboard(world["wb"])


@when(parsers.parse('je passe FlowLoadMode à "{mode}" dans la feuille Config'))
def _switch_mode(world, mode):
    set_config(world["wb"], "FlowLoadMode", mode)
    world["wb"].Application.CalculateFull()


# --- then --------------------------------------------------------------------------------------------

@then("aucune cellule des feuilles de flux ne montre une erreur Excel")
def _no_errors(world):
    assert sheets.error_cells(world["wb"]) == {}


@then("la feuille StatusIntervals reproduit les intervalles de la référence Python")
def _intervals(world):
    excel = {}
    for row in table_dicts(world["wb"], "StatusIntervals"):
        excel.setdefault(row["Key"], []).append(row)
    assert len(excel) == len(world["items"]) == 250
    wrong = []
    for item in world["items"]:
        mine = excel[item.key]
        if len(mine) != len(item.intervals):
            wrong.append((item.key, "count", len(mine), len(item.intervals)))
            continue
        for got, want in zip(mine, item.intervals):
            ok = (got["Status"] == want.status and got["Category"] == want.category and got["Seq"] == want.seq
                  and same_moment(got["EnteredUtc"], want.entered) and same_moment(got["LeftUtc"], want.left)
                  and abs(got["DurationDays"] - want.duration_days) < TOLERANCE and got["IsOpen"] is want.is_open)
            if not ok:
                wrong.append((item.key, want.seq))
    assert wrong == [], wrong[:5]


@then("la table ItemFlow reproduit les calculs de la référence Python")
def _itemflow(world):
    excel = {row["Key"]: row for row in table_dicts(world["wb"], "ItemFlow")}
    assert set(excel) == {i.key for i in world["items"]}
    wrong = []
    for item in world["items"]:
        row = excel[item.key]
        checks = {
            "Team": row["Team"] == item.team,
            "WorkTypeBucket": row["WorkTypeBucket"] == (item.work_type if item.work_type in BUCKETS else "Unmapped"),
            "StartUtc": same_moment(row["StartUtc"], item.start),
            "FirstStartUtc": same_moment(row["FirstStartUtc"], item.first_start),
            "DoneUtc": same_moment(row["DoneUtc"], item.done),
            "RemovedUtc": same_moment(row["RemovedUtc"], item.removed),
            "FlowTimeDays": same_number(row["FlowTimeDays"], item.flow_time_days),
            "LeadTimeDays": same_number(row["LeadTimeDays"], item.lead_time_days),
            "ActiveDays": same_number(row["ActiveDays"], item.active_days),
            "WaitDays": same_number(row["WaitDays"], item.wait_days),
            "UnclassifiedDays": same_number(row["UnclassifiedDays"], item.unclassified_days),
            "BacklogReturns": row["BacklogReturns"] == item.backlog_returns,
            "Reopened": row["Reopened"] == item.reopened,
            "SkippedStart": row["SkippedStart"] == ("Y" if item.skipped_start else "N"),
            "Excluded": row["Excluded"] == (1 if item.excluded_from_flow_time else 0),
        }
        wrong += [(item.key, column, row[column]) for column, ok in checks.items() if not ok]
    assert wrong == [], wrong[:8]
    assert sum(1 for i in world["items"] if i.done) > 20 and any(i.backlog_returns and i.start for i in world["items"])


@then(parsers.parse('la table DailyLoad reproduit la référence Python en mode "{mode}"'))
def _daily_load(world, mode):
    as_of, items = world["dataset"].as_of, world["items"]
    first, end = world["periods"][0].start, world["periods"][-1].end
    expected = {team: daily_load(items, mode, team, first, end, as_of) for team in {i.team for i in items}}
    rows = table_dicts(world["wb"], "DailyLoad")
    keys = [(row["Team"], dt.datetime(1899, 12, 30) + dt.timedelta(days=row["Date"])) for row in rows]
    assert len(keys) == len(set(keys)), "a (team, day) appears twice"
    assert {team for team, _day in keys} >= set(expected), "a team is missing from DailyLoad"
    checked = 0
    for row, (team, day) in zip(rows, keys):
        mine = expected.get(team, {})
        if row["Load"] == "":
            assert day not in mine, (team, day)                       # no snapshot for a day that has not finished
        else:
            assert row["Load"] == mine.get(day, 0), (team, day, row["Load"], mine.get(day, 0))
            checked += 1
    for team, loads in expected.items():                              # every reference figure is in the sheet
        assert all(table_load == loads[day] for (t, day), table_load in
                   ((k, r["Load"]) for k, r in zip(keys, rows) if k[0] == team and k[1] in loads))
        assert {day for t, day in keys if t == team and rows[keys.index((t, day))]["Load"] != ""} == set(loads), team
    assert checked > 500


NA = -2146826246                                         # Excel #N/A as read by Value2
REFERENCE_FIELD = {"Velocity": "velocity", "FlowTimeMedian": "flow_time_median", "FlowLoadAverage": "flow_load_average",
                   "FlowEfficiency": "flow_efficiency", "FlowPredictability": "flow_predictability"}


def chart_blocks(wb):
    """{chart name: (series names, period labels, rows of points)} read from the sheet ChartData."""
    ws, periods = wb.Worksheets("ChartData"), wb.Worksheets("Periods").UsedRange.Rows.Count - 1
    blocks = {}
    for index, spec in enumerate(D.CHARTS):
        top, count = D.block_top(index, periods), len(D.series_names(spec) or team_names(wb))
        grid = ws.Range(ws.Cells(top + 1, 1), ws.Cells(top + 1 + periods, 1 + count)).Value2
        blocks[spec.name] = ([grid[0][c] for c in range(1, count + 1)], [grid[r][0] for r in range(1, periods + 1)],
                             [list(grid[r][1:]) for r in range(1, periods + 1)])
    return blocks


def team_names(wb):
    return [team for team, _formula in sheets._team_rows(wb)]


def expected_point(spec, period, team, expected, metrics_rows):
    """Reference figure as drawn: a number, or NA for « not computable ». Where the reference has no row (a team
    without any issue) the figure is the one of the Metrics table (checked against the reference elsewhere)."""
    want = expected.get((period, team))
    value = getattr(want, REFERENCE_FIELD[spec.column]) if want else metrics_rows[(period, team)][spec.column]
    return NA if isinstance(value, str) else value


def same_point(got, wanted) -> bool:
    if wanted == NA:
        return got == NA
    return not isinstance(got, str) and got != NA and abs(got - wanted) < TOLERANCE


def expected_distribution(spec, period, selected, expected, bucket):
    rows = [m for (p, t), m in expected.items() if p == period and (selected == D.ALL_TEAMS or t == selected)]
    if selected == D.ALL_TEAMS:
        total = sum(m.velocity for m in rows)
        return NA if total == 0 else sum(m.distribution_done[bucket] * m.velocity for m in rows
                                         if m.distribution_done != NOT_COMPUTABLE) / total
    if not rows or rows[0].distribution_done == NOT_COMPUTABLE:
        return NA
    return rows[0].distribution_done[bucket]


@then(parsers.parse('la feuille ChartData reproduit les métriques de la référence Python en mode "{mode}", '
                    "avec #N/A pour « not computable »"))
def _chart_data(world, mode):
    wb, as_of = world["wb"], world["dataset"].as_of
    expected = {(m.period, m.team): m for m in flow_metrics(world["items"], world["periods"], mode, as_of)}
    metrics_rows = {(r["Period"], r["Team"]): r for r in table_dicts(wb, "Metrics")}
    blocks, wrong, gaps = chart_blocks(wb), [], 0
    for spec in D.CHARTS:
        names, labels, rows = blocks[spec.name]
        if spec.column is None:
            continue
        assert names == team_names(wb) and labels == [p.name for p in world["periods"]]
        for label, row in zip(labels, rows):
            for team, got in zip(names, row):
                wanted = expected_point(spec, label, team, expected, metrics_rows)
                gaps += wanted == NA
                if not same_point(got, wanted):
                    wrong.append((spec.name, label, team, got, wanted))
    assert wrong == [], wrong[:8]
    assert gaps > 20                                    # not computable figures are gaps, and they exist in this data
    assert all(got == NA for _l, row in zip(*blocks["flow_predictability"][1:]) for got in row)


@then("chaque graphique a une série par équipe ou par type de travail avec les valeurs de ChartData")
def _series(world):
    wb, dash = world["wb"], world["wb"].Worksheets("Dashboard")
    blocks = chart_blocks(wb)
    assert {co.Name for co in dash.ChartObjects()} == {s.name for s in D.CHARTS}
    for spec in D.CHARTS:
        names, _labels, rows = blocks[spec.name]
        chart = dash.ChartObjects(spec.name).Chart
        assert chart.SeriesCollection().Count == len(names), spec.name
        assert chart.ChartType == spec.chart_type or chart.SeriesCollection(1).ChartType == spec.chart_type, spec.name
        for offset, name in enumerate(names, start=1):
            series = chart.SeriesCollection(offset)
            assert series.Name == name, (spec.name, series.Name, name)
            drawn = [None if row[offset - 1] == NA else row[offset - 1] for row in rows]   # a gap is read as None
            assert list(series.Values) == drawn, (spec.name, name)
        assert chart.HasTitle and chart.ChartTitle.Text == spec.title
    note = dash.Cells(*(lambda a: (a[0] + D.ROWS_PER_CHART, a[1]))(D.chart_anchor(5))).Value2
    assert "not computable" in note and "PI objectives" in note, note


@then("le graphique Flow Distribution suit l'équipe choisie dans la cellule B2")
def _distribution(world):
    wb, dash, as_of = world["wb"], world["wb"].Worksheets("Dashboard"), world["dataset"].as_of
    expected = {(m.period, m.team): m for m in flow_metrics(world["items"], world["periods"], "A", as_of)}
    by_team = {}
    for (_p, team), m in expected.items():
        by_team[team] = by_team.get(team, 0) + m.velocity
    teams = [D.ALL_TEAMS, max(by_team, key=by_team.get), "ART"]
    assert dash.Range(D.SELECTOR_CELL).Value2 == D.ALL_TEAMS                         # the default
    seen = set()
    for selected in teams:
        dash.Range(D.SELECTOR_CELL).Value = com.as_cell(selected)
        wb.Application.Calculate()
        names, labels, rows = chart_blocks(wb)["flow_distribution"]
        assert names == list(BUCKETS)
        for label, row in zip(labels, rows):
            for bucket, got in zip(names, row):
                wanted = expected_distribution(None, label, selected, expected, bucket)
                assert same_point(got, wanted), (selected, label, bucket, got, wanted)
                seen.add((selected, wanted == NA))
        chart = dash.ChartObjects("flow_distribution").Chart
        assert list(chart.SeriesCollection(1).Values) == [None if r[0] == NA else r[0] for r in rows]   # follows B2
    assert {s for s, _na in seen} == set(teams) and any(not na for _s, na in seen)
    dash.Range(D.SELECTOR_CELL).Value = com.as_cell(D.ALL_TEAMS)
    wb.Application.Calculate()


@then("les graphiques sont exportés en PNG")
def _png(world):
    files = sheets.export_charts(world["wb"], OUT_DIR / "dashboard")
    assert set(files) == {s.name for s in D.CHARTS}
    for path in files.values():
        data = open(path, "rb").read()
        assert data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) > 5000, path


@then(parsers.parse('la table Metrics reproduit la référence Python en mode "{mode}"'))
def _metrics(world, mode):
    as_of = world["dataset"].as_of
    expected = {(m.period, m.team): m for m in flow_metrics(world["items"], world["periods"], mode, as_of)}
    rows = table_dicts(world["wb"], "Metrics")
    seen, wrong = 0, []
    for row in rows:
        want = expected.get((row["Period"], row["Team"]))
        if want is None:
            assert row["Velocity"] == 0, row                         # a team without any issue
            continue
        seen += 1
        pairs = {
            "Velocity": (row["Velocity"], want.velocity), "FlowTimeCount": (row["FlowTimeCount"], want.flow_time_count),
            "FlowTimeExcluded": (row["FlowTimeExcluded"], want.flow_time_excluded),
            "FlowTimeMedian": (row["FlowTimeMedian"], want.flow_time_median),
            "FlowTimeP85": (row["FlowTimeP85"], want.flow_time_p85),
            "FlowLoadAverage": (row["FlowLoadAverage"], want.flow_load_average),
            "FlowLoadEnd": (row["FlowLoadEnd"], want.flow_load_end),
            "FlowEfficiency": (row["FlowEfficiency"], want.flow_efficiency),
            "FlowPredictability": (row["FlowPredictability"], want.flow_predictability),
        }
        for bucket in BUCKETS:
            for column, shares in (("DistDone", want.distribution_done), ("DistWip", want.distribution_wip)):
                pairs[f"{column}{bucket}"] = (row[f"{column}{bucket}"],
                                              NOT_COMPUTABLE if shares == NOT_COMPUTABLE else shares[bucket])
        for column, (got, wanted) in pairs.items():
            if column == "FlowPredictability":
                ok = got == wanted == NOT_COMPUTABLE
            else:
                ok = same_figure(got, wanted)
            if not ok:
                wrong.append((row["Period"], row["Team"], column, got, wanted))
    assert wrong == [], wrong[:8]
    assert seen == len(expected) > 50
    figures = [r["FlowTimeMedian"] for r in rows if r["FlowTimeMedian"] != NOT_COMPUTABLE]
    assert len(figures) > 10 and any(r["FlowEfficiency"] != NOT_COMPUTABLE for r in rows)
