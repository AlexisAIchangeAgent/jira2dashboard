"""The user manual of the delivered workbook (pure text; flow/sheets.py writes it)."""
from jira2dashboard import manual
from jira2dashboard.flow import blueprint, dashboard_spec as D
from jira2dashboard.params import model

MANDATORY = model.mandatory_parameters(next(s for s in model.load_all() if s.name == "Config"))


def line_with(lines, text):
    return next(n for n, line in enumerate(lines) if text in line)


def test_the_first_steps_click_read_the_configuration_before_extract():
    steps = manual.first_steps()
    assert steps[0].startswith("First time") and manual.MANUAL_SHEET in steps[0]
    assert line_with(steps, "Sheet Config") < line_with(steps, D.CONFIG_CAPTION) < line_with(steps, D.EXTRACT_CAPTION)
    assert D.CANCEL_CAPTION in steps[line_with(steps, D.EXTRACT_CAPTION)]


def test_the_first_steps_name_the_mandatory_values_typed_at_the_client():
    step = manual.first_steps()[1]
    assert all(name in step for name in model.TYPED_AT_CLIENT)
    assert "ApiToken" in manual.first_steps()[-1]                 # and say to empty the token before sending


def test_the_manual_lists_the_mandatory_parameters_it_is_given():
    rows = manual.manual(MANDATORY)
    assert f"The mandatory parameters: {', '.join(MANDATORY)}." in [text for _kind, text in rows]
    assert MANDATORY == ("BaseUrl", "Email", "ApiToken", "FlowLoadMode", "AsOf", "FeatureTypes", "BlocksLinkType",
                         "FillRateThreshold")


def test_the_manual_has_a_heading_per_section_and_one_per_button():
    headings = [text for kind, text in manual.manual(MANDATORY) if kind == manual.HEADING]
    assert headings[0] == "jira2dashboard - user manual"
    assert f'Button "{D.CONFIG_CAPTION}"' in headings and f'Button "{D.EXTRACT_CAPTION}"' in headings
    assert {kind for kind, _text in manual.manual(MANDATORY)} == {manual.HEADING, manual.TEXT}


def test_the_manual_names_only_sheets_of_the_workbook():
    text = " ".join(line for _kind, line in manual.manual(MANDATORY)) + " ".join(manual.first_steps())
    sheets = {spec.name for spec in model.SPECS} | {blueprint.SHEET_FORMULAS, blueprint.SHEET_CHARTS,
                                                     blueprint.SHEET_CHECKS}
    named = {"Config", "Projects", "Periods", "PIs", "Workflows", "Transitions", "FlowFormulas", "ChartSpec",
             "ChecksSpec"}
    assert named <= sheets
    assert all(name in text for name in named)
