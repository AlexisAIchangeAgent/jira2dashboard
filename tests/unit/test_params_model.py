"""Parameters: load params\\*.md, type the cells, validate. Pure Python, no Excel."""
import shutil

import pytest

from jira2dashboard.params import model
from jira2dashboard.params.model import ParamsError
from jira2dashboard.paths import PARAMS_DIR


@pytest.fixture
def params_copy(tmp_path):
    """A writable copy of the real params folder."""
    target = tmp_path / "params"
    shutil.copytree(PARAMS_DIR, target)
    return target


def edit(folder, filename, old, new, count=1):
    path = folder / filename
    text = path.read_text(encoding="utf-8")
    assert old in text, old
    path.write_text(text.replace(old, new, count), encoding="utf-8")


def sheet(sheets, name):
    return next(s for s in sheets if s.name == name)


# --- the real parameters of the repo ---------------------------------------------------------------

def test_real_parameters_load_in_the_order_of_the_workbook_sheets():
    sheets = model.load_all()
    assert [s.name for s in sheets] == ["Config", "Projects", "Periods", "PIs", "Workflows", "Transitions", "FieldMapping"]


def test_real_parameters_have_the_documented_sizes():
    sheets = model.load_all()
    assert len(sheet(sheets, "Workflows").rows) == 20        # 10 statuses x 2 workflows
    assert len(sheet(sheets, "Transitions").rows) == 26      # 13 moves x 2 workflows
    assert len(sheet(sheets, "FieldMapping").rows) == 22     # 22 columns of the extract


def test_workflows_cells_are_typed():
    rows = sheet(model.load_all(), "Workflows").rows
    assert rows[0] == ("Team", 1, "Nouveau", "Backlog", 0, 0, 0)
    assert rows[9] == ("Team", None, "Dépriorisé", "Backlog", 0, 0, 1)      # "-" means no order


def test_every_row_has_as_many_cells_as_the_header():
    for s in model.load_all():
        assert all(len(r) == len(s.header) for r in s.rows), s.name


# --- validation: clear messages, all problems at once -----------------------------------------------

def test_unknown_category_is_refused(params_copy):
    edit(params_copy, "workflows.md", "| Team | 2 | En analyse | Active |", "| Team | 2 | En analyse | Actif |")
    with pytest.raises(ParamsError, match=r"Workflows row 2.*Category 'Actif'"):
        model.load_all(params_copy)


def test_flag_must_be_0_or_1(params_copy):
    edit(params_copy, "workflows.md", "| Active | 1 | 0 | 0 |", "| Active | 2 | 0 | 0 |")
    with pytest.raises(ParamsError, match="IsStart"):
        model.load_all(params_copy)


def test_order_must_be_a_whole_number_or_a_dash(params_copy):
    edit(params_copy, "workflows.md", "| Team | 3 | Prêt", "| Team | abc | Prêt")
    with pytest.raises(ParamsError, match="Order.*whole number.*'abc'"):
        model.load_all(params_copy)


def test_each_workflow_needs_exactly_one_start_status(params_copy):
    edit(params_copy, "workflows.md", "| Team | 3 | Prêt | Wait | 0 |", "| Team | 3 | Prêt | Wait | 1 |")
    with pytest.raises(ParamsError, match="Team.*exactly one IsStart"):
        model.load_all(params_copy)


def test_a_status_cannot_appear_twice_in_a_workflow(params_copy):
    edit(params_copy, "workflows.md", "| Team | 4 | Planifié |", "| Team | 4 | Prêt |")
    with pytest.raises(ParamsError, match="Team.*'Prêt'.*twice"):
        model.load_all(params_copy)


def test_unknown_workflow_name_is_refused(params_copy):
    edit(params_copy, "workflows.md", "| Team | 1 |", "| Equipe | 1 |")
    with pytest.raises(ParamsError, match="Workflow 'Equipe'"):
        model.load_all(params_copy)


def test_transition_to_an_unknown_status_is_refused(params_copy):
    edit(params_copy, "transitions.md", "| Team | Nouveau | En analyse |", "| Team | Nouveau | En analyze |")
    with pytest.raises(ParamsError, match=r"Transitions row 1.*'En analyze'.*Team"):
        model.load_all(params_copy)


def test_duplicate_transition_is_refused(params_copy):
    edit(params_copy, "transitions.md", "| Team | Prêt | En analyse | back |", "| Team | Nouveau | En analyse | back |")
    with pytest.raises(ParamsError, match="twice"):
        model.load_all(params_copy)


def test_field_role_must_be_known_but_may_have_a_suffix(params_copy):
    edit(params_copy, "field_mapping.md", "| Required for PI views |", "| Mandatory |")
    with pytest.raises(ParamsError, match="Role 'Mandatory'"):
        model.load_all(params_copy)


def test_can_be_empty_is_yes_or_no(params_copy):
    edit(params_copy, "field_mapping.md", "| text | no | readable", "| text | maybe | readable")
    with pytest.raises(ParamsError, match="CanBeEmpty 'maybe'"):
        model.load_all(params_copy)


def test_wrong_header_is_refused(params_copy):
    edit(params_copy, "workflows.md", "| Workflow | Order |", "| Flow | Order |")
    with pytest.raises(ParamsError, match="workflows.md.*header"):
        model.load_all(params_copy)


def test_missing_file_is_refused_with_its_name(params_copy):
    (params_copy / "transitions.md").unlink()
    with pytest.raises(ParamsError, match="transitions.md"):
        model.load_all(params_copy)


def test_a_file_without_a_table_is_refused(params_copy):
    (params_copy / "workflows.md").write_text("# nothing here\n", encoding="utf-8")
    with pytest.raises(ParamsError, match="workflows.md: no table found"):
        model.load_all(params_copy)


def test_empty_status_and_empty_transition_type_are_refused(params_copy):
    edit(params_copy, "workflows.md", "| Team | 3 | Prêt |", "| Team | 3 |  |")
    edit(params_copy, "transitions.md", "| forward |", "|  |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "Workflows row 3: Status is empty" in str(err.value)
    assert "Transitions row 1: Type is empty" in str(err.value)


def test_transition_with_an_unknown_workflow_is_refused(params_copy):
    edit(params_copy, "transitions.md", "| Team | Nouveau | En analyse |", "| Equipe | Nouveau | En analyse |")
    with pytest.raises(ParamsError, match="Transitions row 1: Workflow 'Equipe'"):
        model.load_all(params_copy)


def test_a_workflow_without_any_status_is_refused(params_copy):
    path = params_copy / "workflows.md"
    lines = [ln for ln in path.read_text(encoding="utf-8").splitlines() if not ln.startswith("| Feature |")]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    with pytest.raises(ParamsError, match="workflow 'Feature' has no status"):
        model.load_all(params_copy)


def test_duplicate_extract_column_and_empty_logical_name_are_refused(params_copy):
    edit(params_copy, "field_mapping.md", "| Cle | Key |", "| Projet |  |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "Logical is empty" in str(err.value) and "column 'Projet' listed twice" in str(err.value)


# --- Projects: which projects to extract, and which one is the Program -------------------------------

def test_projects_default_to_the_script_keys_with_ttt_as_the_program():
    rows = sheet(model.load_all(), "Projects").rows
    assert [r[0] for r in rows] == ["TTT", "AAA", "EEE", "FFFF", "TTTT", "SSSS", "SSS", "TETET", "CCCC"]
    assert [r[0] for r in rows if r[1] == "Program"] == ["TTT"]
    assert all(r[2] for r in rows if r[1] == "Team") and rows[0][2] is None     # teams are named, the program is not


def test_project_keys_put_the_program_first_whatever_the_row_order(params_copy):
    path = params_copy / "projects.md"
    text = path.read_text(encoding="utf-8")
    path.write_text(text.replace("| TTT | Program |  |\n", "").replace("| CCCC | Team | CCCC |\n",
                                 "| CCCC | Team | CCCC |\n| TTT | Program |  |\n"), encoding="utf-8")
    assert model.project_keys(model.load_all(params_copy))[0] == "TTT"
    assert model.project_keys()[0] == "TTT"


def test_exactly_one_program_is_needed(params_copy):
    edit(params_copy, "projects.md", "| AAA | Team | AAA |", "| AAA | Program |  |")
    with pytest.raises(ParamsError, match="exactly one Program project is needed, found 2"):
        model.load_all(params_copy)
    edit(params_copy, "projects.md", "| TTT | Program |  |", "| TTT | Team | TTT |")
    edit(params_copy, "projects.md", "| AAA | Program |  |", "| AAA | Team | AAA |")
    with pytest.raises(ParamsError, match="found 0"):
        model.load_all(params_copy)


def test_project_rows_are_checked(params_copy):
    edit(params_copy, "projects.md", "| EEE | Team | EEE |", "| ee-e | Chef |  |")
    edit(params_copy, "projects.md", "| FFFF | Team | FFFF |", "| AAA | Team | Again |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    text = str(err.value)
    assert "'ee-e' is not a Jira project key" in text and "Role 'Chef'" in text and "project 'AAA' listed twice" in text


def test_a_team_project_needs_a_team_name(params_copy):
    edit(params_copy, "projects.md", "| SSS | Team | SSS |", "| SSS | Team |  |")
    with pytest.raises(ParamsError, match="needs a Team name"):
        model.load_all(params_copy)


# --- Config: the connection parameters, proxy included, live in the workbook ------------------------

def test_config_has_the_connection_and_proxy_parameters_with_safe_defaults():
    config = {row[0]: row[1] for row in sheet(model.load_all(), "Config").rows}
    assert list(config) == list(model.CONFIG_PARAMETERS)
    assert config["ProxyWindowsAuth"] == "Yes" and config["TimeoutSeconds"] == "60" and config["MaxRetries"] == "4"
    for client_value in ("BaseUrl", "Email", "ApiToken", "Filter", "WorkTypeField", "ProxyServer", "ProxyBypass",
                         "StoryPointsField", "BusinessValueField", "PIField", "TeamField", "SprintField",
                         "TargetStartField", "TargetEndField"):
        assert config[client_value] is None, f"{client_value} must stay empty in the repo"


def test_data_check_parameters_have_defaults_and_are_checked(params_copy):
    config = {row[0]: row[1] for row in sheet(model.load_all(), "Config").rows}
    assert (config["FeatureTypes"], config["SubTaskTypes"], config["PointsRequiredTypes"], config["BlocksLinkType"],
            config["FillRateThreshold"]) == ("Feature", "Sub-task", "Story", "Blocks", "90")
    edit(params_copy, "config.md", "| FillRateThreshold | 90 |", "| FillRateThreshold | 120 |")
    with pytest.raises(ParamsError, match="FillRateThreshold must be a whole number between 0 and 100"):
        model.load_all(params_copy)


def test_pis_are_empty_in_the_repo_and_checked_like_periods(params_copy):
    assert sheet(model.load_all(), "PIs").rows == ()
    path = params_copy / "pis.md"
    rows = ["| PI#01-2026 | 2026-01-05 | 2026-03-30 |", "| PI#01-2026 | 2026-04-01 | tomorrow |",
            "| PI#02-2026 | 2026-06-01 | 2026-05-01 |"]
    path.write_text(path.read_text(encoding="utf-8") + "\n".join(rows) + "\n", encoding="utf-8")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    text = str(err.value)
    assert "pi 'PI#01-2026' listed twice" in text and "PIs row 2: Start and End must be dates" in text
    assert "PIs row 3: Start must be before End" in text


def test_flow_parameters_default_to_mode_a_and_no_fixed_as_of_date():
    config = {row[0]: row[1] for row in sheet(model.load_all(), "Config").rows}
    assert config["FlowLoadMode"] == "A" and config["AsOf"] is None


def test_log_level_defaults_to_steps():
    config = {row[0]: row[1] for row in sheet(model.load_all(), "Config").rows}
    assert config["LogLevel"] == "Steps"


def test_log_level_must_be_steps_or_errors(params_copy):
    edit(params_copy, "config.md", "| LogLevel | Steps |", "| LogLevel | Debug |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "LogLevel must be Steps or Errors" in str(err.value)


def test_flow_load_mode_and_as_of_are_checked(params_copy):
    edit(params_copy, "config.md", "| FlowLoadMode | A |", "| FlowLoadMode | C |")
    edit(params_copy, "config.md", "| AsOf |  |", "| AsOf | tomorrow |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "FlowLoadMode must be A or B" in str(err.value) and "AsOf must be a UTC date" in str(err.value)


@pytest.mark.parametrize("moment", ["2026-11-03", "2026-11-03 18:00", "2026-11-03T18:00:00"])
def test_as_of_accepts_a_date_or_a_date_and_time(params_copy, moment):
    edit(params_copy, "config.md", "| AsOf |  |", f"| AsOf | {moment} |")
    config = {row[0]: row[1] for row in sheet(model.load_all(params_copy), "Config").rows}
    assert config["AsOf"] == moment


# --- Periods: the iterations, empty in the repo -----------------------------------------------------

def test_periods_are_empty_in_the_repo_because_the_calendar_is_the_clients():
    periods = sheet(model.load_all(), "Periods")
    assert periods.header == ("Period", "Start", "End") and periods.rows == ()


def test_period_rows_are_checked(params_copy):
    path = params_copy / "periods.md"
    path.write_text(path.read_text(encoding="utf-8") + "| I1 | 2026-10-05 | 2026-10-19 |\n| I1 | 2026-10-19 | soon |\n"
                    "| I3 | 2026-11-02 | 2026-10-26 |\n|  | 2026-11-02 | 2026-11-16 |\n", encoding="utf-8")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    text = str(err.value)
    assert "period 'I1' listed twice" in text and "Periods row 2: Start and End must be dates" in text
    assert "Periods row 3: Start must be before End" in text and "Periods row 4: Period needs a name" in text


def test_valid_period_rows_are_accepted(params_copy):
    path = params_copy / "periods.md"
    path.write_text(path.read_text(encoding="utf-8") + "| I1 | 2026-10-05 | 2026-10-19 |\n| I2 | 2026-10-19 | 2026-11-02 |\n",
                    encoding="utf-8")
    assert len(sheet(model.load_all(params_copy), "Periods").rows) == 2


def test_an_api_token_in_the_repo_is_refused(params_copy):
    edit(params_copy, "config.md", "| ApiToken |  |", "| ApiToken | abc123 |")
    with pytest.raises(ParamsError, match="ApiToken must stay empty"):
        model.load_all(params_copy)


def test_config_values_are_checked(params_copy):
    edit(params_copy, "config.md", "| ProxyWindowsAuth | Yes |", "| ProxyWindowsAuth | Maybe |")
    edit(params_copy, "config.md", "| TimeoutSeconds | 60 |", "| TimeoutSeconds | soon |")
    edit(params_copy, "config.md", "| BaseUrl |  |", "| BaseUrl | digitalplatform |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    text = str(err.value)
    assert "ProxyWindowsAuth must be Yes or No" in text and "TimeoutSeconds must be a whole number" in text
    assert "BaseUrl must start with http" in text


def test_config_parameter_names_are_checked(params_copy):
    edit(params_copy, "config.md", "| ProxyBypass |", "| ProxyBypas |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "unknown parameter 'ProxyBypas'" in str(err.value) and "'ProxyBypass' is missing" in str(err.value)


def test_all_problems_are_reported_together(params_copy):
    edit(params_copy, "workflows.md", "| Team | 2 | En analyse | Active |", "| Team | 2 | En analyse | Actif |")
    edit(params_copy, "field_mapping.md", "| Required for PI views |", "| Mandatory |")
    with pytest.raises(ParamsError) as err:
        model.load_all(params_copy)
    assert "Actif" in str(err.value) and "Mandatory" in str(err.value)
