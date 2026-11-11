"""The workbook to deliver, without Excel: which modules it holds, which macro its button runs, where it goes."""
from jira2dashboard import build_workbook as bw
from jira2dashboard.paths import INTER_DIR, RELEASE_DIR
from jira2dashboard.testbench import run_vba_tests as rvt


def test_the_delivered_workbook_has_no_test_module():
    assert not {"modTests", "modButtons", "modBackground"} & set(bw.MODULES)
    assert "modCoverage" in bw.MODULES                  # the other modules call Cov
    assert set(bw.MODULES) <= set(rvt.ORDER)
    assert [m for m in rvt.ORDER if m in bw.MODULES] == bw.MODULES      # same order: a module comes after what it uses


def test_the_button_runs_a_public_sub_without_arguments_of_a_delivered_module():
    module, macro = bw.BUTTON_MACRO.split(".")
    assert module in bw.MODULES
    text = rvt.module_sources([module])[module]
    assert f"Public Sub {macro}()" in text


def test_every_module_of_the_delivery_is_ascii():
    assert set(rvt.module_sources(bw.MODULES)) == set(bw.MODULES)


def test_the_macro_of_the_button_calls_the_chain_the_scenario_tests():
    text = rvt.module_sources(["modActions"])["modActions"]
    for call in ("ExtractFlowData", "BuildFlowSheetsIn ThisWorkbook"):
        assert call in text


def test_the_workbook_to_deliver_goes_to_out_release():
    assert bw.TARGET == bw.target_for(bw.DELIVERY_NAME) == RELEASE_DIR / "jira2dashboard.xlsm"


def test_a_variant_named_by_out_goes_to_out_inter():
    assert bw.target_for("try.xlsm") == INTER_DIR / "try.xlsm"


def test_the_version_of_a_release_is_given_to_the_build():
    assert bw.parse(["--version", "v1.0.0"]).version == "v1.0.0"
    assert bw.parse([]).version is None
