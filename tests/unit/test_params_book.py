"""The parameters of a filled workbook, read as a file by the Python extraction (no Excel)."""
import datetime as dt

import pytest

from jira2dashboard import mock_workbook as M
from jira2dashboard.flow.checks import CHECKS
from jira2dashboard.flow.model import ACTIVE
from jira2dashboard.mockjira.generator import AS_OF
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN
from jira2dashboard.params.model import ParamsError, load_all
from jira2dashboard.standalone import params_book as P
from jira2dashboard.standalone.xlsx_write import Sheet, write_xlsx

URL = "http://127.0.0.1:8099"


def book(tmp_path, values=None, change=None):
    """A parameters workbook of the fake Jira; `change(sheets)` edits the sheets before they are written."""
    sheets = M.params_book(URL, values)
    if change:
        change(sheets)
    return write_xlsx(tmp_path / "params.xlsx", sheets)


def named(sheets, name):
    return next(s for s in sheets if s.name == name)


def replace(sheets, name, new):
    index = next(n for n, s in enumerate(sheets) if s.name == name)
    sheets[index] = new


def test_the_sheets_come_back_in_the_form_of_the_md_files(tmp_path):
    sheets = P.read_params(book(tmp_path))
    assert [s.name for s in sheets] == list(P.SHEETS)
    config = {row[0]: row[1] for row in named(sheets, "Config").rows}
    assert (config["BaseUrl"], config["Email"], config["ApiToken"], config["AsOf"]) == (
        URL, MOCK_EMAIL, MOCK_TOKEN, "2026-11-03 18:00")
    assert named(sheets, "Iterations").rows[0] == ("I1", "2026-06-01", "2026-06-15")       # dates typed in Excel: serials
    assert named(sheets, "PIs").rows == (("PI#01-2026", "2026-06-01", "2026-12-31"),)
    md = {s.name: s for s in load_all()}
    for name in ("Projects", "Workflows", "Transitions"):
        assert named(sheets, name).rows == md[name].rows, name


def test_the_settings_of_a_run(tmp_path):
    settings = P.run_settings(P.read_params(book(tmp_path)))
    c = settings.connection
    assert (c.base_url, c.email, c.token, c.proxy_server, c.windows_auth, c.timeout_s, c.max_retries) == (
        URL, MOCK_EMAIL, MOCK_TOKEN, "", True, 60, 4)
    assert settings.as_of == AS_OF and settings.mode == "A" and settings.iterations == M.iterations()
    assert settings.fields["StoryPointsField"] == "customfield_10016" and settings.blocks_link_type == "Blocks"
    assert settings.rules.rule("Team", "En analyse").category == ACTIVE
    assert settings.checks.ready(next(c for c in CHECKS if c.id == "DC18"))
    assert settings.projects[0][1] == "Program"


def test_an_empty_mandatory_value_stops_with_the_message_of_the_vba(tmp_path):
    sheets = P.read_params(book(tmp_path, {"Email": "", "AsOf": ""}))
    with pytest.raises(ParamsError, match=r"^Fill these values of the Config sheet first \(Mandatory = True\): "
                                          r"Email, AsOf$"):
        P.run_settings(sheets)


def test_a_workbook_built_before_the_columns_example_and_mandatory_is_read(tmp_path):
    def old(sheets):
        config = named(sheets, "Config")
        keep = [config.header.index(c) for c in ("Parameter", "Value", "Description")]
        replace(sheets, "Config", Sheet("Config", tuple(config.header[i] for i in keep),
                                        [tuple(row[i] for i in keep) for row in config.rows]))
    sheets = P.read_params(book(tmp_path, change=old))
    config = named(sheets, "Config")
    assert config.header == ("Parameter", "Value", "Example", "Mandatory", "Description")
    assert P.missing_mandatory(config) == [] and P.run_settings(sheets).as_of == AS_OF


def test_a_flag_typed_as_text_boolean_or_decimal_stays_a_number(tmp_path):
    def typed(sheets):
        workflows = named(sheets, "Workflows")
        rows = [list(r) for r in workflows.rows]
        rows[0][4], rows[1][4], rows[2][4] = "0", False, 0.0          # Nouveau, En analyse (the start), Prêt: IsStart
        rows[1][4] = True
        replace(sheets, "Workflows", Sheet("Workflows", workflows.header, [tuple(r) for r in rows]))
    sheets = P.read_params(book(tmp_path, change=typed))
    rows = named(sheets, "Workflows").rows
    assert (rows[0][4], rows[1][4], rows[2][4]) == (0, 1, 0)
    rules = P.run_settings(sheets).rules
    assert not rules.rule("Team", "Nouveau").is_start and rules.rule("Team", "En analyse").is_start


def test_as_of_typed_as_a_date_in_excel_is_read(tmp_path):
    sheets = P.read_params(book(tmp_path, {"AsOf": dt.datetime(2026, 11, 3, 18, 0)}))
    assert P.run_settings(sheets).as_of == dt.datetime(2026, 11, 3, 18, 0)


def test_the_defaults_and_the_readings_of_the_vba(tmp_path):
    def connection(values):
        return P.run_settings(P.read_params(book(tmp_path, values))).connection
    c = connection({"TimeoutSeconds": "", "MaxRetries": "", "ProxyWindowsAuth": ""})
    assert (c.timeout_s, c.max_retries, c.windows_auth) == (60, 4, True)
    c = connection({"TimeoutSeconds": "0", "MaxRetries": "-2", "ProxyWindowsAuth": "oui"})
    assert (c.timeout_s, c.max_retries, c.windows_auth) == (60, 0, True)
    c = connection({"TimeoutSeconds": 30, "MaxRetries": "abc", "ProxyWindowsAuth": "No"})
    assert (c.timeout_s, c.max_retries, c.windows_auth) == (30, 0, False)
    assert P.run_settings(P.read_params(book(tmp_path, {"FlowLoadMode": " b "}))).mode == "B"


@pytest.mark.parametrize("values, message", [
    ({"FlowLoadMode": "C"}, "FlowLoadMode must be A or B"),
    ({"FillRateThreshold": "120"}, "FillRateThreshold must be a whole number between 0 and 100"),
    ({"AsOf": "3 novembre"}, "AsOf must be a UTC date"),
    ({"BaseUrl": "https://me:tok@site.atlassian.net"}, "must not contain a user name"),
])
def test_a_config_value_the_vba_could_not_use_is_refused(tmp_path, values, message):
    with pytest.raises(ParamsError, match=message):
        P.read_params(book(tmp_path, values))


def test_the_rules_of_the_md_files_apply_to_the_sheets(tmp_path):
    def wrong(sheets):
        workflows = named(sheets, "Workflows")
        rows = [list(r) for r in workflows.rows]
        rows[1][3] = "Actif"
        replace(sheets, "Workflows", Sheet("Workflows", workflows.header, [tuple(r) for r in rows]))
    with pytest.raises(ParamsError, match=r"Workflows row 2: Category 'Actif'"):
        P.read_params(book(tmp_path, change=wrong))


def test_a_missing_sheet_or_column_is_named(tmp_path):
    def broken(sheets):
        sheets.remove(named(sheets, "PIs"))
        projects = named(sheets, "Projects")
        replace(sheets, "Projects", Sheet("Projects", ("Project", "Team"), [(r[0], r[2]) for r in projects.rows]))
    with pytest.raises(ParamsError) as error:
        P.read_params(book(tmp_path, change=broken))
    assert "The workbook has no sheet PIs" in str(error.value)
    assert "Sheet Projects: no column Role in its first row" in str(error.value)


def test_no_iteration_stops_the_run(tmp_path):
    def empty(sheets):
        replace(sheets, "Iterations", Sheet("Iterations", ("Iteration-label", "Start", "End"), []))
    with pytest.raises(ParamsError, match="Sheet Iterations has no iteration"):
        P.run_settings(P.read_params(book(tmp_path, change=empty)))


def test_a_cell_is_read_like_a_pasted_value():
    assert [P.text(v) for v in (None, "", "  x\t\r\n", True, False, 60.0, 2.5, 7)] == [
        None, None, "x", "True", "False", "60", "2.5", "7"]
    assert P.serial_to_iso(46329.75) == "2026-11-03 18:00:00" and P.serial_to_iso(46174) == "2026-06-01"
