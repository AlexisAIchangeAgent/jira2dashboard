"""Characterization tests: pure helpers of testbench/run_vba_tests.py (no Excel needed)."""
import urllib.error
import urllib.request

import pytest

from jira2dashboard.mockjira.server import PAC_PATH, PAC_TYPE, pac_script
from jira2dashboard.testbench import run_vba_tests as rvt


def test_module_sources_reads_and_normalizes(tmp_path, monkeypatch):
    (tmp_path / "modA.bas").write_text('Attribute VB_Name = "modA"\nOption Explicit\n', encoding="ascii")
    monkeypatch.setattr(rvt, "VBA_DIR", tmp_path)
    assert rvt.module_sources(["modA"]) == {"modA": "Option Explicit\r\n"}


def test_module_sources_refuses_non_ascii_source(tmp_path, monkeypatch):
    (tmp_path / "modA.bas").write_bytes("Sub A() ' café\n".encode("utf-8"))
    monkeypatch.setattr(rvt, "VBA_DIR", tmp_path)
    with pytest.raises(UnicodeDecodeError):
        rvt.module_sources(["modA"])


def test_coverage_markers_keeps_order_and_drops_duplicates():
    sources = {"m1": 'Cov "B1"\nCov "A1"\n', "m2": 'x = 1\nCov "B1"\nCov "C1"\n'}
    assert rvt.coverage_markers(sources) == ["B1", "A1", "C1"]


def test_coverage_markers_ignores_other_text():
    assert rvt.coverage_markers({"m": 'Coverage "X"\nCov X\n'}) == []


def test_the_local_server_serves_one_proxy_script_for_t_proxy_script():
    with rvt.pac_server() as url:
        assert url.startswith("http://127.0.0.1:") and url.endswith(PAC_PATH)
        with urllib.request.urlopen(url, timeout=5) as response:
            assert response.headers["Content-Type"] == PAC_TYPE
            assert response.read().decode("ascii") == pac_script(rvt.PAC_PROXY)
        with pytest.raises(urllib.error.HTTPError, match="404"):
            urllib.request.urlopen(url.replace(PAC_PATH, "/other"), timeout=5)
    with pytest.raises(OSError):                                    # closed once the tests are done
        urllib.request.urlopen(url, timeout=2)


def test_real_vba_modules_are_ascii_and_have_unique_markers():
    sources = rvt.module_sources(rvt.ORDER)
    assert len(sources) == 16
    markers = rvt.coverage_markers(sources)
    # 65 originally + modParams PRM01..PRM23 + TXT16 (TrimWhite) + JIR14..JIR21 (proxy, Config sheet, hints, filter)
    # + modExtract MEX01..MEX09 (the job: MEX07-09), JIR23 (search page)
    assert len(markers) == 293          # + TXT17 (malformed Jira date), PRM24 (AsOf typed as a date), JIR22 (http BaseUrl refused), FLB30-32, FLW08, MEX04-06
    #                                     + code review 2026-11-04: FLB35, HTP08-09, HTP10-11 (Cancel during a retry wait), TXT18-28 (UTF-8 and Base64 in plain VBA), JIR24-25, JSN19-20, PRM25
    #                                     + modLog LOG01..LOG31 (the local log of the runs, its level and run flag, the hint and Notepad after an error, the support codes)
    #                                     + LOG32..LOG37 (the inputs of the run in PARAM lines, the code on RUN end stopped)
    #                                     + LOG38..LOG40 (dates in ISO, the version of the workbook), HTP12 (RETRY line with no answer)
    #                                     - the folder rule of the Rovo CSV (3 PRM markers; PRM53-55 renumbered PRM50-52)
    #                                     + PRM53 (UNKNOWN apart from the empty values), TXT29-30 (the status as rows)
    #                                     + FLB37-40 (the PI axis of the Feature Predictability chart, the table PIMetrics)
    #                                     + modJiraConfig JCF01..JCF34 (the Jira configuration read into proposal sheets)
    #                                     + PRM26..PRM29 (the column Mandatory of the Config sheet)
    #                                     + modProxy PXY01..PXY11, JIR26-27 (the proxy of the Windows settings)
    #                                     + MEX10..MEX13 (the raw sheets in steps: issue rows, intervals, write, default batch)
    assert sorted(m for m in markers if m.startswith("LOG")) == [f"LOG{n:02d}" for n in range(1, 41)]
    assert sorted(m for m in markers if m.startswith("JCF")) == [f"JCF{n:02d}" for n in range(1, 35)]
    assert sorted(m for m in markers if m.startswith("PRM")) == [f"PRM{n:02d}" for n in range(1, 54)]
    assert sorted(m for m in markers if m.startswith("MEX")) == [f"MEX{n:02d}" for n in range(1, 14)]
    assert sorted(m for m in markers if m.startswith("PXY")) == [f"PXY{n:02d}" for n in range(1, 12)]
    # + modFlowBuild FLB01..FLB39 (build of the flow sheets, the data checks and the charts, the PI axis)
    assert sorted(m for m in markers if m.startswith("FLB")) == [f"FLB{n:02d}" for n in range(1, 41)]


def test_every_button_macro_exists_in_modbuttons():
    text = rvt.module_sources(["modButtons"])["modButtons"]
    for _cell, _caption, macro in rvt.BUTTONS:
        assert f"Sub {macro.split('.')[1]}" in text
