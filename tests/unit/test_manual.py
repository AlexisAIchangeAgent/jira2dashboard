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
    assert MANDATORY == ("BaseUrl", "Email", "ApiToken", "FlowLoadMode", "PredictabilityMode", "AsOf", "FeatureTypes",
                         "BlocksLinkType", "FillRateThreshold")


def test_the_manual_has_a_heading_per_section_and_one_per_button():
    headings = [text for kind, text in manual.manual(MANDATORY) if kind == manual.HEADING]
    assert headings[0] == "jira2dashboard - user manual"
    assert f'Button "{D.CONFIG_CAPTION}"' in headings and f'Button "{D.EXTRACT_CAPTION}"' in headings
    assert f'Button "{D.COMPUTE_CAPTION}"' in headings and f'Button "{D.DRAW_CAPTION}"' in headings
    assert {kind for kind, _text in manual.manual(MANDATORY)} == {manual.HEADING, manual.TEXT}


def test_the_manual_names_only_sheets_of_the_workbook():
    text = " ".join(line for _kind, line in manual.manual(MANDATORY)) + " ".join(manual.first_steps())
    sheets = {spec.name for spec in model.SPECS} | {blueprint.SHEET_FORMULAS, blueprint.SHEET_CHARTS,
                                                     blueprint.SHEET_CHECKS}
    named = {"Config", "Projects", "Iterations", "PIs", "Workflows", "Transitions", "FlowFormulas", "ChartSpec",
             "ChecksSpec"}
    assert named <= sheets
    assert all(name in text for name in named)


def test_the_buttons_have_their_section_in_the_order_of_the_dashboard():
    headings = [text for kind, text in manual.manual(MANDATORY) if kind == manual.HEADING]
    captions = (D.ROVO_CAPTION, D.ROVO_LOAD_CAPTION, D.CONFIG_CAPTION, D.PROPOSE_CAPTION, D.EXTRACT_CAPTION, D.COMPUTE_CAPTION, D.DRAW_CAPTION,
                D.ROVO_DEBUG_CAPTION)
    at = [headings.index(f'Button "{caption}"') for caption in captions]
    assert at == sorted(at)


def test_the_manual_says_how_to_check_the_figures_with_the_counts_of_the_code():
    """AL, 2026-11-24: the guide to check the metrics and their data, in the sheet Manual; the first steps point to it."""
    from jira2dashboard import sheet_names as S
    from jira2dashboard.flow.checks import CHECKS
    from jira2dashboard.flow.reconciliation import CONTROLS
    rows = manual.manual(MANDATORY)
    start = rows.index((manual.HEADING, manual.CHECK_SECTION))
    end = next(n for n in range(start + 1, len(rows)) if rows[n][0] == manual.HEADING)
    section = " ".join(text for _kind, text in rows[start + 1:end])
    assert f"{len(CONTROLS)} controls" in section and "Gap must be 0" in section
    for name in ("Reconciliation", "VelocityCheck", "FlowTimeCheck", "FlowEfficiencyCheck", "FlowLoadCheck", "FlowDistCheck", "MetricItems",
                 "StatusRules", "Workflows", "ItemFlow"):
        assert S.of(name) in section, name
    assert f"{len(CHECKS)} data checks" in " ".join(text for _kind, text in rows)
    assert manual.CHECK_SECTION in manual.first_steps()[6]


def test_the_manual_stays_short():
    rows = manual.manual(MANDATORY)
    assert max(len(text) for _kind, text in rows) <= 420 and sum(len(text) for _kind, text in rows) <= 10_500      # AL, 2026-11-26: + the button Debug prompt for Rovo
    assert sum(len(line) for line in manual.first_steps()) <= 1_300
