"""Steps of features/flow_excel.feature (the Excel formulas against the Python reference, cell by cell), of
features/jira_config.feature (the button that reads the Jira configuration into proposal sheets) and of
features/first_use.feature (the first steps, the sheet Manual and the column Mandatory of Config)."""
import dataclasses
import datetime as dt
import gc
import os
import time
from pathlib import Path

import pytest
import pywintypes
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import manual, mock_workbook
from jira2dashboard.build_workbook import build_delivery
from jira2dashboard.excel import com
from jira2dashboard.flow import blueprint, sheets
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import daily_load
from jira2dashboard.flow.metrics import BUCKETS, flow_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Rules
from jira2dashboard.flow.checks import CHECKS, IDS, PER_ISSUE, CheckSettings, Moves, fill_rates, run_checks
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW, with_anomalies
from jira2dashboard.mockjira import config as jira_config
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MockJira
from jira2dashboard.params.model import load_all, mandatory_parameters, project_keys
from jira2dashboard.paths import INTER_DIR
from jira2dashboard.standalone.results import data_check_rows

scenarios("../features/flow_excel.feature")
scenarios("../features/jira_config.feature")             # same workbook and fake Jira steps
scenarios("../features/first_use.feature")

TARGET = INTER_DIR / "flow_workbook.xlsm"
TOLERANCE = 1e-6                    # days (0.09 s): Excel dates are doubles, the reference works in whole seconds


@pytest.fixture(scope="module")
def built_workbook():
    """The workbook to deliver, built once per module by the same code as `run.ps1 build` (no test module in it)."""
    build_delivery(TARGET)
    return TARGET


def excel_scope(fixture_name, config):
    """A fresh Excel per scenario; J2D_ONE_EXCEL=1 shares one Excel across the module (the stress setup of the open defect, CLAUDE.md)."""
    return "module" if os.environ.get("J2D_ONE_EXCEL") == "1" else "function"


@pytest.fixture(scope=excel_scope)
def flow_workbook(built_workbook):
    """A fresh Excel per scenario, on the workbook as built: each scenario stands alone. One Excel for the whole module
    (several extractions and builds in one process) crashed in about 2 runs out of 5 (open defect, CLAUDE.md)."""
    holder = {}
    with com.excel_session() as app:
        app.AutomationSecurity = com.MSO_AUTOMATION_SECURITY_LOW
        holder["wb"] = app.Workbooks.Open(str(built_workbook.resolve()))
        code = holder["wb"].VBProject.VBComponents(holder["wb"].CodeName).CodeModule
        assert "AbortExtraction" in code.Lines(1, code.CountOfLines)       # closing the workbook drops a pending slice
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


@given(parsers.parse("un faux Jira lent avec {n:d} tickets générés avec les projets des paramètres"))
def _slow_jira(world, jira_server, n):
    world["dataset"] = generate(seed=11, issues=n, projects=project_keys())
    world["jira"] = jira_server(world["dataset"], credentials=("me@example.com", "tok-123"), delay=0.15, page_clamp=5)


@given(parsers.parse("un faux Jira avec {n:d} tickets générés et les anomalies de données connues"))
def _jira_with_anomalies(world, jira_server, n):
    dataset, world["anomalies"] = with_anomalies(generate(seed=11, issues=n, projects=project_keys()), project_keys())
    world["dataset"] = dataset
    world["jira"] = jira_server(dataset, credentials=("me@example.com", "tok-123"))


@given(parsers.parse('le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "{mode}"'))
def _config(world, flow_workbook, mode):
    wb = world["wb"] = flow_workbook["wb"]
    for name, value in {"BaseUrl": world["jira"].base_url, "Email": "me@example.com", "ApiToken": "tok-123",
                        "WorkTypeField": world["dataset"].fields.work_type, "FlowLoadMode": mode,
                        "AsOf": f"{world['dataset'].as_of:%Y-%m-%d %H:%M}"}.items():
        set_config(wb, name, value)


@given(parsers.parse("un premier calcul sur un faux Jira de {n:d} tickets a déjà rempli le classeur"))
def _first_run(world, jira_server, n):
    """The workbook holds the results of a bigger extraction: the next one must replace them and leave nothing over."""
    wb, first = world["wb"], generate(seed=11, issues=n, projects=project_keys())
    server = jira_server(first, credentials=("me@example.com", "tok-123"))
    set_config(wb, "BaseUrl", server.base_url)
    set_config(wb, "AsOf", f"{first.as_of:%Y-%m-%d %H:%M}")
    wb.Application.Run(f"'{wb.Name}'!modActions.ExtractAndCompute")
    assert len(table_dicts(wb, "Issues")) == n
    set_config(wb, "BaseUrl", world["jira"].base_url)
    set_config(wb, "AsOf", f"{world['dataset'].as_of:%Y-%m-%d %H:%M}")


FIELD_IDS = mock_workbook.FIELD_IDS          # Config parameter -> attribute of dataset.fields


def put_pis(wb, rows):
    ws = wb.Worksheets("PIs")
    ws.Range("A2:C20").ClearContents()
    for n, (name, start, end) in enumerate(rows, start=2):
        ws.Cells(n, 1).Value = com.as_cell(name)
        ws.Cells(n, 2).Value2 = com.to_excel_serial(start)
        ws.Cells(n, 3).Value2 = com.to_excel_serial(end)


@given("les champs Jira des contrôles et la fenêtre du PI sont renseignés dans les feuilles Config et PIs")
def _check_fields(world):
    ids = {name: getattr(world["dataset"].fields, attr) for name, attr in FIELD_IDS.items()}
    for name, value in ids.items():
        set_config(world["wb"], name, value)
    put_pis(world["wb"], [(PI, *PI_WINDOW)])
    world["settings"] = dataclasses.replace(CheckSettings.from_params(), fields=ids | {}, pis={PI: PI_WINDOW})


@given("les champs Jira des contrôles et la feuille PIs sont vides")
def _check_fields_empty(world):
    for name in FIELD_IDS:
        set_config(world["wb"], name, "")
    put_pis(world["wb"], [])
    world["settings"] = dataclasses.replace(CheckSettings.from_params(), fields={n: "" for n in FIELD_IDS}, pis={})


@given(parsers.parse('les itérations de 14 jours depuis le "{first}" jusqu\'à la date de référence dans la feuille Iterations'))
def _iterations(world, first):
    iterations = mock_workbook.iterations(dt.datetime.fromisoformat(first), world["dataset"].as_of)
    mock_workbook.write_iterations(world["wb"].Worksheets("Iterations"), iterations)    # the iterations of an earlier run go
    world["iterations"] = iterations


# --- when --------------------------------------------------------------------------------------------

@when("j'exécute la macro Extraire et calculer du classeur")
def _extract_and_compute(world):
    wb = world["wb"]
    wb.Application.Run(f"'{wb.Name}'!modActions.ExtractAndCompute")
    reference(world)


def reference(world):
    """The Python reference of the data the fake Jira holds, for the steps that compare with the workbook."""
    rules, as_of = Rules.from_params(), world["dataset"].as_of
    world["rules"] = rules
    world["items"] = [item_flow(i, as_of, rules) for i in to_flow_issues(world["dataset"])]


BUSY = (-2147418111, -2147417846)               # Excel is running a macro slice: the call is rejected, try again


def retry(call):
    """Run a COM call between two slices of the extraction: Excel refuses calls while a slice runs."""
    last = None
    for _ in range(400):
        try:
            return call()
        except AttributeError as err:                                    # a busy Excel also fails the lookup of a name
            last = err
            time.sleep(0.02)
        except pywintypes.com_error as err:
            if err.hresult not in BUSY:
                raise
            last = err
            time.sleep(0.02)
    raise TimeoutError("Excel stayed busy") from last                    # the real error stays visible


def put(dash, value):
    dash.Range("R1").Value = value


def macro(world, name):
    """Application.Run of a modActions macro. The Application proxy and the workbook name are read once per scenario:
    each wb.Application is one more COM call (and a type lookup) for an Excel that is running the job."""
    if "app" not in world:
        wb = world["wb"]
        world["app"], world["prefix"] = wb.Application, f"'{wb.Name}'!modActions."
    app, prefix = world["app"], world["prefix"]
    return retry(lambda: app.Run(prefix + name))


def status_and_bar(world):
    """(status text, bar as a fraction of its width) in one call: between two slices Excel answers only a few calls."""
    bar, _sep, text = macro(world, "ProgressSnapshot").partition("|")
    return text, float(bar)


POLL_S = 0.25        # a human pace: fewer COM calls into the Excel that runs the job


def wait_for(world, done, seconds=180):
    """Poll the status cell and the bar until done(status) is true; returns the (status, bar) samples."""
    samples, deadline = [], time.time() + seconds
    while time.time() < deadline:
        try:
            samples.append(status_and_bar(world))
        except Exception as err:                              # noqa: BLE001 - say where the run was when Excel went away
            raise RuntimeError(f"{err!r}; the last samples: {samples[-4:]} (Excel gone? check the EXCEL.EXE exit code: "
                               "CLAUDE.md, open defect)") from err
        if done(samples[-1][0]):
            return samples
        time.sleep(POLL_S)
    raise TimeoutError(samples[-3:])


@when("je clique sur le bouton Extraire et calculer sans attendre la fin")
def _click(world, request):
    request.addfinalizer(lambda: macro(world, "AbortExtraction"))        # no slice must outlive the scenario
    world["before"] = {name: len(table_dicts(world["wb"], name)) for name in ("Issues", "ItemFlow")}         if "items" in world else None
    world["log_from"] = LOG_FILE.stat().st_size if LOG_FILE.exists() else 0
    macro(world, "BtnExtractAndCompute")
    reference(world)


@when("j'annule dès que la barre a avancé")
def _cancel(world):
    world["samples"] = wait_for(world, lambda text: text.startswith("Reading") and "page 2" in text or text.startswith("Done"))
    assert not world["samples"][-1][0].startswith("Done"), "the extraction ended before the cancel"
    macro(world, "BtnCancelExtraction")
    world["samples"] += wait_for(world, lambda text: text.startswith(("Cancelled", "Stopped", "Done")))


@then("la barre de progression avance sans reculer, le classeur répond entre deux tranches et l'extraction se termine")
def _progress(world):
    wb, dash = world["wb"], world["wb"].Worksheets("Dashboard")
    ticks = checks = 0

    def usable(text):
        nonlocal ticks, checks
        ticks += 1
        if ticks % 10 == 1:                                              # a cell can be written while the job runs
            checks += 1
            retry(lambda: put(dash, ticks))
            assert retry(lambda: dash.Range("R1").Value2) == ticks
        return text.startswith(("Done", "Stopped"))

    samples = wait_for(world, usable)
    assert checks >= 3, f"the workbook was written only {checks} times during the run"
    bars = [bar for _text, bar in samples]
    assert samples[-1][0].startswith("Done"), samples[-1]
    assert bars == sorted(bars) and bars[-1] > 0.99, bars[-5:]
    assert len({round(b, 2) for b in bars}) >= 4, "the bar did not move step by step"
    kinds = {text.split(":")[0] for text, _bar in samples}
    assert "Reading Jira" in kinds, kinds        # the search has many slices; the other steps are one slice each (not always seen)
    dash.Range("R1").ClearContents()
    assert retry(lambda: wb.Worksheets("ItemFlow").ListObjects("ItemFlow").ListRows.Count) == len(world["items"])


LOG_FILE = Path(os.environ["LOCALAPPDATA"]) / "jira2dashboard" / "logs" / "jira2dashboard.log"      # modLog.LogFile


@then("le journal local décrit l'exécution étape par étape, sans jeton, adresse ni réponse de Jira")
def _log(world):
    raw = LOG_FILE.read_bytes()
    start = world["log_from"] if len(raw) >= world["log_from"] else 0                 # a new file when the log was big
    text = raw[start:].decode("ascii")                                                  # plain ASCII, or this fails
    lines = [line.split(" ", 3)[2:] for line in text.splitlines()                      # date time KIND text
             if " RUN previous run did not end " not in line]                           # an earlier scenario crashed
    kinds = [kind for kind, *_ in lines]
    assert lines[0][1].startswith("button Extract from Jira and compute | Excel "), lines[0]
    assert lines[-1] == ["RUN", "end ok"], lines[-1]
    steps = kinds[1:-1]
    assert steps == ["BEGIN", "END"] * (len(steps) // 2), kinds                          # each step ends before the next
    begun = [rest[0] for kind, *rest in lines if kind == "BEGIN"]
    for name in ("search page 1", "changelogs from issue 1 of 150", "raw sheets: write", "flow sheets: charts"):
        assert name in begun, (name, begun[:5])
    for secret in ("tok-123", "me@example.com", "127.0.0.1", "http", "{"):
        assert secret not in text, secret


@then("les résultats du premier calcul sont conservés et l'état dit « Cancelled »")
def _cancelled(world):
    text, bar = world["samples"][-1]
    assert text.startswith("Cancelled"), text
    assert {name: len(table_dicts(world["wb"], name)) for name in ("Issues", "ItemFlow")} == world["before"]
    assert bar == pytest.approx(0, abs=0.01)


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
    assert len(excel) == len(world["items"]) > 50
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


def expected_flags(world):
    """{key: ids} from the Python reference, with the field ids and PIs that the workbook was given."""
    return run_checks(to_flow_issues(world["dataset"]), world["dataset"].as_of, world["rules"], Moves.from_params(),
                      world["settings"])


def excel_flags(world):
    return {row["Key"]: {c for c in PER_ISSUE if row[c] == 1} for row in table_dicts(world["wb"], "ItemFlow")}


@then("la table DataChecks et les colonnes DC de ItemFlow reproduisent les contrôles de la référence Python")
def _data_checks(world):
    expected, settings = expected_flags(world), world["settings"]
    flags = excel_flags(world)
    assert flags == {key: set(ids) for key, ids in expected.items()}
    assert all(row[c] in (0, 1) for row in table_dicts(world["wb"], "ItemFlow") for c in PER_ISSUE)
    rows = {row["Id"]: row for row in table_dicts(world["wb"], "DataChecks")}
    assert list(rows) == list(IDS)                                         # same ids, same order as the plan
    reference_rows = data_check_rows(expected, world["fill_rates"], settings)     # also what the Python path writes
    for _id, title, severity, *_needs, _source, _ready, items, result in reference_rows:
        row = rows[_id]
        wanted = (title, severity, "" if items is None else items, result)
        assert (row["Title"], row["Severity"], row["Items"], row["Result"]) == wanted, _id


@then("la table FillRates reproduit les taux de remplissage de la référence Python")
def _fill_rates(world):
    dataset, settings = world["dataset"], world["settings"]
    issues = to_flow_issues(dataset)
    teams = sorted({i.team for i in issues})
    types = sorted({i.type for i in issues})
    got = {(r["Field"], r["Team"], r["Type"]): (r["Count"], r["Filled"], r["Below"]) for r in table_dicts(world["wb"], "FillRates")}
    reference = fill_rates(issues, settings, teams, types)
    world["fill_rates"] = reference
    expected = {(r.field, r.team, r.type): (r.count, r.filled, 1 if r.below else 0) for r in reference}
    assert got == expected


@then("chaque anomalie connue est trouvée par les contrôles qui lui correspondent et seulement par eux")
def _known_anomalies(world):
    flags = excel_flags(world)
    assert {key: flags[key] for key in world["anomalies"]} == {k: set(v) for k, v in world["anomalies"].items()}


@then("les contrôles qui lisent un champ non configuré ne signalent aucun ticket")
def _unconfigured(world):
    unready = {c.id for c in CHECKS if not world["settings"].ready(c)}
    assert unready and {"DC11", "DC17"} <= unready
    assert not {c for ids in excel_flags(world).values() for c in ids} & unready


@then(parsers.parse('la table DailyLoad reproduit la référence Python en mode "{mode}"'))
def _daily_load(world, mode):
    as_of, items = world["dataset"].as_of, world["items"]
    first, end = world["iterations"][0].start, world["iterations"][-1].end
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
    assert checked > 300


NA = -2146826246                                         # Excel #N/A as read by Value2
REFERENCE_FIELD = {"Velocity": "velocity", "FlowTimeMedian": "flow_time_median", "FlowTimeP85": "flow_time_p85", "FlowLoadAverage": "flow_load_average",
                   "FlowEfficiency": "flow_efficiency", "FlowPredictability": "flow_predictability"}


def chart_blocks(wb, world):
    """{chart name: (series names, row keys, rows of points)} read from the sheet ChartData. A row is an iteration, or a day
    for the daily curve; the series of a band are not in the names (they are checked on their own)."""
    ws, iterations = wb.Worksheets("ChartData"), len(world["iterations"])
    days = (world["iterations"][-1].end - world["iterations"][0].start).days
    blocks = {}
    for index, spec in enumerate(D.CHARTS):
        n = days if spec.axis == D.DAY_AXIS else iterations
        top, count = D.block_top(index, iterations), len(D.series_names(spec) or team_names())
        grid = ws.Range(ws.Cells(top + 1, 1), ws.Cells(top + 1 + n, 1 + count + (2 if spec.band else 0))).Value2
        blocks[spec.name] = ([grid[0][c] for c in range(1, count + 1)], [grid[r][0] for r in range(1, n + 1)],
                             [list(grid[r][1:count + 1]) for r in range(1, n + 1)])
        if spec.band:
            blocks[spec.name + ".band"] = ([grid[0][count + 1], grid[0][count + 2]], blocks[spec.name][1],
                                           [list(grid[r][count + 1:]) for r in range(1, n + 1)])
    return blocks


def team_names():
    projects = next(s for s in load_all() if s.name == "Projects")
    return blueprint.team_names([tuple(row) for row in projects.rows])


@then("les feuilles de flux ont exactement la taille des nouvelles données")
def _sizes(world):
    wb, teams, iterations = world["wb"], team_names(), world["iterations"]
    size = {name: len(table_dicts(wb, name)) for name in ("Issues", "StatusIntervals", "ItemFlow", "DailyLoad", "Metrics")}
    assert size["Issues"] == size["ItemFlow"] == len(world["items"]) == 120
    assert size["StatusIntervals"] == sum(len(item.intervals) for item in world["items"])
    assert size["DailyLoad"] == (iterations[-1].end - iterations[0].start).days * len(teams)
    assert size["Metrics"] == len(iterations) * len(teams)
    assert wb.Worksheets("Dashboard").ChartObjects().Count == len(D.CHARTS) == 8
    names = [sheet.Name for sheet in wb.Worksheets]
    assert len(names) == len(set(names)), names


def expected_point(spec, iteration, team, expected, metrics_rows):
    """Reference figure as drawn: a number, or NA for « not computable ». Where the reference has no row (a team
    without any issue) the figure is the one of the Metrics table (checked against the reference elsewhere)."""
    want = expected.get((iteration, team))
    value = getattr(want, REFERENCE_FIELD[spec.column]) if want else metrics_rows[(iteration, team)][spec.column]
    return NA if isinstance(value, str) else value


def same_point(got, wanted) -> bool:
    if wanted == NA:
        return got == NA
    return not isinstance(got, str) and got != NA and abs(got - wanted) < TOLERANCE


def expected_distribution(iteration, selected, expected, bucket):
    rows = [m for (p, t), m in expected.items() if p == iteration and (selected == D.ALL_TEAMS or t == selected)]
    if selected == D.ALL_TEAMS:
        total = sum(m.velocity for m in rows)
        return NA if total == 0 else sum(m.distribution_done[bucket] * m.velocity for m in rows
                                         if m.distribution_done != NOT_COMPUTABLE) / total
    if not rows or rows[0].distribution_done == NOT_COMPUTABLE:
        return NA
    return rows[0].distribution_done[bucket]


def daily_curve_differences(world, mode, block):
    """The daily Flow Load curve against the reference: the load at the end of each day, a gap after the as-of date."""
    names, days, rows = block
    as_of, items = world["dataset"].as_of, world["items"]
    first, end = world["iterations"][0].start, world["iterations"][-1].end
    reference = {team: daily_load(items, mode, team, first, end, as_of) for team in {i.team for i in items}}
    wrong = []
    for serial, row in zip(days, rows):
        day = dt.datetime(1899, 12, 30) + dt.timedelta(days=serial)
        for team, got in zip(names, row):
            wanted = reference.get(team, {}).get(day, 0) if day + dt.timedelta(days=1) <= as_of else NA
            if not same_point(got, wanted):
                wrong.append((day.date(), team, got, wanted))
    assert any(got != NA for row in rows for got in row) and any(got == NA for row in rows for got in row)
    return wrong


@then(parsers.parse('la feuille ChartData reproduit les métriques de la référence Python en mode "{mode}", '
                    "avec #N/A pour « not computable »"))
def _chart_data(world, mode):
    wb, as_of = world["wb"], world["dataset"].as_of
    expected = {(m.iteration, m.team): m for m in flow_metrics(world["items"], world["iterations"], mode, as_of)}
    metrics_rows = {(r["Iteration-label"], r["Team"]): r for r in table_dicts(wb, "Metrics")}
    blocks, wrong, gaps = chart_blocks(wb, world), [], 0
    for spec in D.CHARTS:
        names, labels, rows = blocks[spec.name]
        if spec.column is None:
            continue
        assert names == team_names() and labels == [p.name for p in world["iterations"]]
        for label, row in zip(labels, rows):
            for team, got in zip(names, row):
                wanted = expected_point(spec, label, team, expected, metrics_rows)
                gaps += wanted == NA
                if not same_point(got, wanted):
                    wrong.append((spec.name, label, team, got, wanted))
    assert wrong == [], wrong[:8]
    assert gaps > 10                                    # not computable figures are gaps, and they exist in this data
    daily_wrong = daily_curve_differences(world, mode, blocks["flow_load_daily"])
    assert daily_wrong == [], daily_wrong[:8]
    band = blocks["flow_predictability.band"]
    assert band[0] == list(D.BAND_NAMES) and all(row == [0.8, 1.0] for row in band[2])
    assert all(got == NA for _l, row in zip(*blocks["flow_predictability"][1:]) for got in row)


@then("chaque graphique a une série par équipe ou par type de travail avec les valeurs de ChartData")
def _series(world):
    wb, dash = world["wb"], world["wb"].Worksheets("Dashboard")
    blocks = chart_blocks(wb, world)
    assert {co.Name for co in dash.ChartObjects()} == {s.name for s in D.CHARTS}
    for spec in D.CHARTS:
        names, _labels, rows = blocks[spec.name]
        chart = dash.ChartObjects(spec.name).Chart
        assert chart.SeriesCollection().Count == len(names) + (2 if spec.band else 0), spec.name
        assert chart.ChartType == spec.chart_type or chart.SeriesCollection(1).ChartType == spec.chart_type, spec.name
        for offset, name in enumerate(names, start=1):
            series = chart.SeriesCollection(offset)
            assert series.Name == name, (spec.name, series.Name, name)
            drawn = [None if row[offset - 1] == NA else row[offset - 1] for row in rows]   # a gap is read as None
            assert list(series.Values) == drawn, (spec.name, name)
        if spec.band:                                    # two constant series after the data: 80 % and 100 %
            for offset, expected in ((len(names) + 1, 0.8), (len(names) + 2, 1.0)):
                assert set(chart.SeriesCollection(offset).Values) == {expected}, (spec.name, offset)
                assert chart.SeriesCollection(offset).Name == D.BAND_NAMES[offset - len(names) - 1]
        assert chart.HasTitle and chart.ChartTitle.Text == spec.title
    note = dash.Cells(*(lambda a: (a[0] + D.ROWS_PER_CHART, a[1]))(D.chart_anchor([c.name for c in D.CHARTS].index("flow_predictability")))).Value2
    assert "not computable" in note and "PI objectives" in note, note


@then("le graphique Flow Distribution suit l'équipe choisie dans la cellule B2")
def _distribution(world):
    wb, dash, as_of = world["wb"], world["wb"].Worksheets("Dashboard"), world["dataset"].as_of
    expected = {(m.iteration, m.team): m for m in flow_metrics(world["items"], world["iterations"], "A", as_of)}
    by_team = {}
    for (_p, team), m in expected.items():
        by_team[team] = by_team.get(team, 0) + m.velocity
    teams = [D.ALL_TEAMS, max(by_team, key=by_team.get), "ART"]
    assert dash.Range(D.SELECTOR_CELL).Value2 == D.ALL_TEAMS                         # the default
    seen = set()
    for selected in teams:
        dash.Range(D.SELECTOR_CELL).Value = com.as_cell(selected)
        wb.Application.Calculate()
        names, labels, rows = chart_blocks(wb, world)["flow_distribution"]
        assert names == list(BUCKETS)
        for label, row in zip(labels, rows):
            for bucket, got in zip(names, row):
                wanted = expected_distribution(label, selected, expected, bucket)
                assert same_point(got, wanted), (selected, label, bucket, got, wanted)
                seen.add((selected, wanted == NA))
        chart = dash.ChartObjects("flow_distribution").Chart
        assert list(chart.SeriesCollection(1).Values) == [None if r[0] == NA else r[0] for r in rows]   # follows B2
    assert {s for s, _na in seen} == set(teams) and any(not na for _s, na in seen)
    dash.Range(D.SELECTOR_CELL).Value = com.as_cell(D.ALL_TEAMS)
    wb.Application.Calculate()


@then("les graphiques sont exportés en PNG")
def _png(world):
    files = sheets.export_charts(world["wb"], INTER_DIR / "dashboard")
    assert set(files) == {s.name for s in D.CHARTS}
    for path in files.values():
        data = Path(path).read_bytes()
        assert data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) > 5000, path


@then(parsers.parse('la table Metrics reproduit la référence Python en mode "{mode}"'))
def _metrics(world, mode):
    as_of = world["dataset"].as_of
    expected = {(m.iteration, m.team): m for m in flow_metrics(world["items"], world["iterations"], mode, as_of)}
    rows = table_dicts(world["wb"], "Metrics")
    seen, wrong = 0, []
    for row in rows:
        want = expected.get((row["Iteration-label"], row["Team"]))
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
                wrong.append((row["Iteration-label"], row["Team"], column, got, wanted))
    assert wrong == [], wrong[:8]
    assert seen == len(expected) > 20
    figures = [r["FlowTimeMedian"] for r in rows if r["FlowTimeMedian"] != NOT_COMPUTABLE]
    assert len(figures) > 5 and any(r["FlowEfficiency"] != NOT_COMPUTABLE for r in rows)


# --- jira_config.feature: the button Read the Jira configuration ----------------------------------------------------

PARAMETER_SHEETS = ("Config", "Projects", "Workflows", "Transitions", "FieldMapping", "Iterations")
PREVIEW = "/rest/api/3/workflows/preview"
MSO_FORM_CONTROL = 8                 # Shape.Type of a Form button


def sheet_rows(wb, name):
    """Rows of a proposal sheet as dicts (header -> value); empty cells as ""."""
    values = wb.Worksheets(name).UsedRange.Value
    header, rows = values[0], values[1:]
    return [{h: ("" if v is None else v) for h, v in zip(header, row)} for row in rows]


def param_rows(sheet):
    return next(s for s in load_all() if s.name == sheet).rows


def number_or_text(value):
    return int(value) if isinstance(value, float) and value.is_integer() else value


@given(parsers.parse("le faux Jira refuse l'aperçu des workflows avec {status:d}"))
def _preview_refused(world, status):
    world["jira"].fail_next("POST", PREVIEW, status, times=100)


@when("je clique sur le bouton Lire la configuration Jira")
def _read_configuration(world):
    wb = world["wb"]
    world["params_before"] = {name: wb.Worksheets(name).UsedRange.Value for name in PARAMETER_SHEETS}
    world["log_from"] = LOG_FILE.stat().st_size if LOG_FILE.exists() else 0
    buttons = {shape.TextFrame.Characters().Text: (shape.OnAction, shape.TopLeftCell.Address)
               for shape in wb.Worksheets("Dashboard").Shapes if shape.Type == MSO_FORM_CONTROL}
    on_action, cell = buttons["Read the Jira configuration"]                  # the button the user clicks is there
    assert on_action.endswith(sheets.CONFIG_MACRO), on_action
    assert cell.replace("$", "") == D.CONFIG_CELL, cell
    macro(world, on_action.rpartition(".")[2])                                 # and runs what its OnAction names
    world["status"] = macro(world, "ProgressSnapshot").partition("|")[2]


@then("l'état du tableau de bord dit que la configuration est lue")
def _configuration_read(world):
    assert world["status"].startswith(f"Jira configuration read: {len(project_keys())} project(s), 2 workflow(s), "), world["status"]
    assert "could not be read" not in world["status"] and "not found" not in world["status"], world["status"]


@then("l'état du tableau de bord dit que les workflows n'ont pas pu être lus")
def _workflows_not_read(world):
    assert f"The workflows could not be read for {len(project_keys())} project(s)" in world["status"], world["status"]


@then("la feuille JiraIssueTypes donne le workflow de chaque type de ticket de chaque projet de la feuille Projects")
def _issue_types(world):
    issues = world["dataset"].issues
    expected = {(key, i.type, i.workflow) for key in project_keys() for i in issues if i.project == key}
    rows = sheet_rows(world["wb"], "JiraIssueTypes")
    assert {(r["Project"], r["IssueType"], r["Workflow"]) for r in rows} == expected
    assert len(rows) == len(expected)
    assert {r["Subtask"] for r in rows if r["IssueType"] == "Sub-task"} == {"Yes"}


@then("la feuille JiraStatuses donne les statuts de chaque type de ticket avec leur catégorie Jira")
def _statuses(world):
    dataset = world["dataset"]
    keys = [p["key"] for p in jira_config.project_search(dataset, {"keys": list(project_keys())})["values"]]  # Jira's order
    expected = [(key, t["name"], s["name"], s["statusCategory"]["key"]) for key in keys
                for t in jira_config.project_statuses(dataset, key) for s in t["statuses"]]
    rows = sheet_rows(world["wb"], "JiraStatuses")
    assert [(r["Project"], r["IssueType"], r["Status"], r["JiraCategory"]) for r in rows] == expected


@then("la feuille JiraWorkflows redonne les statuts de la feuille Workflows, avec une catégorie proposée et les statuts libres")
def _workflows(world):
    proposed = {"Backlog": "Backlog", "Active": "", "Wait": "", "Done": "Done", "Removed": "Done"}  # Jira has 3 categories
    expected = {(wf, status): (1 if order == 1 else "", proposed[category], free)
                for wf, order, status, category, _start, _end, free in param_rows("Workflows")}
    rows = sheet_rows(world["wb"], "JiraWorkflows")
    got = {(r["Workflow"], r["Status"]): (number_or_text(r["Order"]), r["Category"], number_or_text(r["IsFree"])) for r in rows}
    assert got == expected
    assert len(rows) == len(expected)
    assert {(r["IsStart"], r["IsEnd"]) for r in rows} == {("", "")}               # a human decision


@then("la feuille JiraTransitions redonne les passages de la feuille Transitions")
def _transitions(world):
    rows = sheet_rows(world["wb"], "JiraTransitions")
    assert sorted((r["Workflow"], r["From"], r["To"]) for r in rows) == sorted(row[:3] for row in param_rows("Transitions"))
    assert {r["Type"] for r in rows} == {""}                                        # forward, back, exit: a human decision


@then("la feuille JiraFields liste les champs du faux Jira avec leur identifiant")
def _fields(world):
    rows = sheet_rows(world["wb"], "JiraFields")
    assert [(r["Id"], r["Name"]) for r in rows] == [(f["id"], f["name"]) for f in world["dataset"].fields.catalog()]


@then("les feuilles de paramètres n'ont pas changé")
def _parameters_kept(world):
    wb = world["wb"]
    assert {name: wb.Worksheets(name).UsedRange.Value for name in PARAMETER_SHEETS} == world["params_before"]


@then("les feuilles JiraWorkflows et JiraTransitions n'ont que leurs en-têtes")
def _no_workflow(world):
    for name in ("JiraWorkflows", "JiraTransitions"):
        assert sheet_rows(world["wb"], name) == [], name


@then("le journal local décrit la lecture de la configuration, sans jeton, adresse ni réponse de Jira")
def _configuration_log(world):
    raw = LOG_FILE.read_bytes()
    start = world["log_from"] if len(raw) >= world["log_from"] else 0
    text = raw[start:].decode("ascii")
    lines = [line.split(" ", 3)[2:] for line in text.splitlines() if " RUN previous run did not end " not in line]
    kinds = [kind for kind, *_ in lines]
    assert lines[0][1].startswith("button Read the Jira configuration | Excel "), lines[0]
    assert lines[-1] == ["RUN", "end ok"], lines[-1]
    assert kinds[1:-1] == ["BEGIN", "END"] * ((len(kinds) - 2) // 2), kinds
    begun = [rest[0] for kind, *rest in lines if kind == "BEGIN"]
    count = len(project_keys())
    assert begun == ["proxy detection", "projects"] + \
        [f"configuration of project {n} of {count}" for n in range(1, count + 1)] + ["fields", "proposal sheets"], begun
    for secret in ("tok-123", "me@example.com", "127.0.0.1", "http", "{", "workflows/preview"):
        assert secret not in text, secret


# --- first use: the first steps, the sheet Manual and the column Mandatory of Config (first_use.feature) --------------

def dashboard_buttons(wb):
    """{caption: OnAction} of the Form buttons of the Dashboard."""
    return {shape.TextFrame.Characters().Text: shape.OnAction
            for shape in wb.Worksheets("Dashboard").Shapes if shape.Type == MSO_FORM_CONTROL}


def column_values(ws, row, column, count):
    return [cells[0] for cells in ws.Range(ws.Cells(row, column), ws.Cells(row + count - 1, column)).Value]


@given("le classeur à livrer")
def _delivered(world, flow_workbook):
    world["wb"] = flow_workbook["wb"]


@then("le Dashboard donne les premières étapes, avec « Read the Jira configuration » avant « Extract from Jira and compute »")
def _first_steps(world):
    dash, steps = world["wb"].Worksheets("Dashboard"), manual.first_steps()
    first = dash.Range(D.STEPS_CELL)
    assert column_values(dash, first.Row, first.Column, len(steps) + 1) == steps + [None]     # nothing after the block
    assert first.Font.Bold
    order = [next(n for n, line in enumerate(steps) if caption in line) for caption in (D.CONFIG_CAPTION, D.EXTRACT_CAPTION)]
    assert order[0] < order[1], order
    assert {D.CONFIG_CAPTION, D.EXTRACT_CAPTION, D.CANCEL_CAPTION} <= set(dashboard_buttons(world["wb"]))


@then("l'onglet Manual suit le Dashboard et reproduit le manuel")
def _manual_sheet(world):
    wb = world["wb"]
    assert [wb.Worksheets(1).Name, wb.Worksheets(2).Name] == ["Dashboard", manual.MANUAL_SHEET]
    config = next(sheet for sheet in load_all() if sheet.name == "Config")
    rows, ws = manual.manual(mandatory_parameters(config)), wb.Worksheets(manual.MANUAL_SHEET)
    assert column_values(ws, 1, 1, len(rows) + 1) == [text for _kind, text in rows] + [None]
    assert [bool(ws.Cells(n, 1).Font.Bold) for n in range(1, len(rows) + 1)] == [kind == manual.HEADING
                                                                                 for kind, _text in rows]


@then("la feuille Config a les colonnes Example et Mandatory")
def _mandatory_column(world):
    ws = world["wb"].Worksheets("Config")
    assert [ws.Cells(1, column).Value for column in range(1, 6)] == [
        "Parameter", "Value", "Example", "Mandatory", "Description"]


@when(parsers.parse('je vide le paramètre "{name}" de la feuille Config'))
def _empty_parameter(world, name):
    set_config(world["wb"], name, "")


@when("je clique sur le bouton Extraire et calculer")
def _click_extract(world, request):
    request.addfinalizer(lambda: macro(world, "AbortExtraction"))        # no slice must outlive the scenario
    on_action = dashboard_buttons(world["wb"])[D.EXTRACT_CAPTION]          # the button the user clicks
    macro(world, on_action.rpartition(".")[2])
    world["status"] = macro(world, "ProgressSnapshot").partition("|")[2]


@then(parsers.parse("l'état du tableau de bord nomme les paramètres obligatoires vides \"{names}\""))
def _mandatory_named(world, names):
    assert world["status"].startswith("Stopped: "), world["status"]
    assert world["status"].endswith(f"(Mandatory = True): {names}"), world["status"]


@then("le faux Jira n'a reçu aucune requête")
def _no_request(world):
    assert world["jira"].requests == []
