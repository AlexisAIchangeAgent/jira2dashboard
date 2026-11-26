"""The texts of the debug prompt for Rovo (rovo_debug.py): what the button writes, checked without Excel (AL, 2026-11-26)."""
import re

from jira2dashboard import rovo_debug as RD
from jira2dashboard import sheet_names as S
from jira2dashboard.flow import reconciliation as R
from jira2dashboard.flow.formulas import metricitems_column
from jira2dashboard.flow.trace import CUTOFF_HOURS
from jira2dashboard.paths import VBA_DIR

TOKEN = re.compile(r"\{(\w+)\}")
# The tokens the VBA fills, per part of the spec (modRovoDebug): a token of a text that the VBA does not know would stay in the file as it is.
KNOWN = {
    RD.HEAD: {"file", "files"},
    RD.SUMMARY: {"version", "as_of", "filter", "feature_types", "projects", "statuses", "iterations_total", "iterations_listed", "window_start",
                 "window_end", "iteration_index", "jql_extraction", "jql_features", "health", "parameters", "controls"},
    RD.ITERATION: {"iteration", "start", "end", "left_out", "cutoff", "teams"},
    RD.TEAM: {"iteration", "team", "velocity", "velocity_reason", "flow_time_count", "done_count", "done_keys", "not_in_flow_time", "jql_missing",
              "jql_extra"},
    RD.JQL: {"program", "feature_types", "projects", "filter", "scope", "filter_jql", "end_statuses", "start", "last", "not_keys", "keys"},
}
SECRETS = ("BaseUrl", "Email", "ApiToken", "ProxyServer", "ProxyBypass", "Authorization", "password")


def texts() -> list[str]:
    return [text for _part, _name, text, hint in RD.spec_rows() for text in (text, hint)]


def test_every_control_has_a_group_a_title_and_a_hint_and_no_data_check_code():
    """Rovo knows no DCnn: the control HISTORIES is named "(DC06, DC07)" in the workbook, so the file prints the title of the spec."""
    rows = [row for row in RD.spec_rows() if row[0] in RD.GROUPS]
    assert [name for _group, name, _title, _hint in rows] == [name for name, _ in R.CONTROLS]
    assert all(title and hint for _group, _name, title, hint in rows)
    assert {group for group, *_ in rows} == set(RD.GROUPS)
    assert not [text for text in texts() if re.search(r"\bDC\d\d\b", text)]
    assert dict((name, group) for group, name, *_ in rows)[R.HISTORIES] == RD.WORKBOOK         # Rovo cannot read a changelog (ROVO-329)


def test_the_controls_checked_in_jira_are_those_a_jql_can_check():
    """The research: only the population, the Velocity and FeatureTypes have a JQL equivalent; the gaps compare the workbook with itself."""
    jira = [name for name, (group, _t, _h) in RD.CONTROL_TEXTS.items() if group == RD.JIRA]
    assert jira == [R.POPULATION, R.VELOCITY_TOTAL, R.VELOCITY_ROWS, R.FEATURE_TYPES]
    assert all(RD.CONTROL_TEXTS[name][0] == RD.WORKBOOK for name in (R.FLOW_TIME_GAP, R.FLOW_EFFICIENCY_GAP, R.FLOW_LOAD_GAP, R.FLOW_DISTRIBUTION_GAP))


def test_every_token_is_one_the_vba_fills():
    source = (VBA_DIR / "modRovoDebug.bas").read_text(encoding="ascii")
    for part, name, text, _hint in RD.spec_rows():
        found = set(TOKEN.findall(text))
        assert found <= KNOWN.get(part, set()), (part, name, found - KNOWN.get(part, set()))
    for part, tokens in KNOWN.items():
        for token in tokens:
            assert f'"{{{token}}}"' in source, (part, token)                # the VBA replaces it: Fill(..., "{token}", value)


def test_the_vba_names_the_team_of_the_program_and_its_items_as_the_reference():
    """modRovoDebug finds the block of the ART by its team name (flow/model.ART, as modHealth); an item that holds ", " is joined with "; " so that
    Rovo can count the items of a list."""
    from jira2dashboard.flow.model import ART
    source = (VBA_DIR / "modRovoDebug.bas").read_text(encoding="ascii")
    assert f'Private Const PROGRAM_TEAM As String = "{ART}"' in source
    assert 'Private Const ITEM_SEP As String = "; "' in source and "JoinWith(why, ITEM_SEP)" in source


def test_the_columns_are_those_of_the_lists_of_metric_items():
    """A reader of a list takes its column from formulas.metricitems_column (CLAUDE.md): a column added to a block moves the next ones."""
    columns = {name: int(text) for part, name, text, _hint in RD.spec_rows() if part == RD.COLUMN}
    assert columns["Done.Key"] == metricitems_column(0, "Key") and columns["Cutoff.Hours"] == metricitems_column(2, "Hours")
    assert len(columns) == sum(len(names) for _block, _number, names in RD.COLUMNS)
    assert columns["LeftOut.Key"] > columns["Done.WhyNotInFlowTime"]                              # the blocks sit side by side


def test_the_limits_and_the_files():
    limits = {name: text for part, name, text, _hint in RD.spec_rows() if part == RD.LIMIT}
    assert limits == {"MaxCharacters": "12000", "MaxIterations": "18"}                            # AL: the last 18 iterations at most
    assert RD.FOLDER == "rovo-debug" and len(S.of(RD.SPEC_SHEET)) <= 31


def test_no_text_names_a_secret_and_every_sheet_by_its_real_name():
    joined = "\n".join(texts())
    assert not [word for word in SECRETS if word.lower() in joined.lower()]
    assert S.of("Reconciliation") in joined and S.of("StatusRules") in joined
    assert "the sheet Reconciliation " not in joined


def test_the_instructions_ask_for_keys_and_causes_never_a_total():
    summary = "\n".join(text for part, _n, text, _h in RD.spec_rows() if part == RD.SUMMARY)
    assert "Run each JQL exactly as it is written" in summary and "Never give a total that you computed" in summary
    assert "how many keys you read in each list" in summary                                     # a file cut by Rovo shows (ROVO-249)
    for cause in ("Time zone", "Created done", "Reopened", "Not a Feature", "Status rules", "Filter or scope", "Rights", "AsOf", "Other"):
        assert f"\n- {cause}: " in summary, cause
    assert "UTC" in summary and f"{CUTOFF_HOURS} hours" in summary


def test_the_jql_are_well_formed_and_leave_out_an_empty_list():
    """An empty `key in ()` is no valid JQL (an iteration where nothing was done): the clause on the keys is a token of its own, left out then."""
    jqls = dict(RD.JQLS)
    for name, text in jqls.items():
        assert text.count("(") == text.count(")") and text.count('"') % 2 == 0, name
        assert "ORDER BY" not in text.upper(), name
    assert "{not_keys}" in jqls["missing"] and "key" not in jqls["missing"].replace("{not_keys}", "")
    assert jqls["not_keys"].startswith(" AND key not in (") and jqls["filter"].startswith(" AND (")
    assert 'DURING ("{start}", "{last}")' in jqls["missing"] and "NOT (status CHANGED TO" in jqls["extra"]
    assert "status in ({end_statuses})" in jqls["missing"]                                    # done now, as the Velocity of the workbook
