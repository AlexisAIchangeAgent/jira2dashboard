"""The workbook to deliver, without Excel: which modules it holds, which macro its button runs, where it goes."""
import pytest

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


# --- the workbook that disappears after Excel saved it (an endpoint protection), seen on another machine ------------------

def stub_the_build(monkeypatch, tmp_path, *, exists_after_build, compile_ok):
    target = tmp_path / "jira2dashboard.xlsm"

    def build(path, release=None):
        if exists_after_build:
            path.write_text("wb")

    monkeypatch.setattr(bw, "target_for", lambda name: target)
    monkeypatch.setattr(bw.clean, "main", lambda: None)
    monkeypatch.setattr(bw, "build_delivery", build)
    monkeypatch.setattr(bw.rvt, "compile_ok", compile_ok)
    return target


def test_a_workbook_gone_right_after_the_save_is_named_with_its_code(monkeypatch, tmp_path, capsys):
    stub_the_build(monkeypatch, tmp_path, exists_after_build=False, compile_ok=lambda t: "COMPILED_OK")
    assert bw.main([]) == 1
    out = capsys.readouterr().out
    assert "J2D-SYS-012" in out and "right after Excel saved it" in out and "endpoint protection" in out
    assert "built " not in out


def test_a_workbook_gone_before_excel_opens_it_is_named_instead_of_a_traceback(monkeypatch, tmp_path, capsys):
    def compile_ok(target):
        target.unlink()                                   # removed between the save and the open of the child Excel
        raise RuntimeError("_check_child.py failed (1): child excel pid 1")
    stub_the_build(monkeypatch, tmp_path, exists_after_build=True, compile_ok=compile_ok)
    assert bw.main([]) == 1
    out = capsys.readouterr().out
    assert "built " in out and "J2D-SYS-012" in out and "when Excel opened it" in out


def test_another_failure_of_the_compile_check_still_raises(monkeypatch, tmp_path):
    def compile_ok(target):
        raise RuntimeError("_check_child.py failed (1): something else")
    stub_the_build(monkeypatch, tmp_path, exists_after_build=True, compile_ok=compile_ok)
    with pytest.raises(RuntimeError, match="something else"):
        bw.main([])


def test_a_good_build_still_returns_0(monkeypatch, tmp_path, capsys):
    stub_the_build(monkeypatch, tmp_path, exists_after_build=True, compile_ok=lambda t: "COMPILED_OK")
    assert bw.main([]) == 0
    assert "compile: COMPILED_OK" in capsys.readouterr().out
