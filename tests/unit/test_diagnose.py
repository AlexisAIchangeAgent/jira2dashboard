"""The pure part of `scripts\\run.ps1 diagnose` (jira2dashboard.diagnose): hashes, origin of the file, masking, log tail."""
import hashlib
import re
from pathlib import Path

import pytest

from jira2dashboard import diagnose as D


def test_fingerprint_gives_md5_sha256_and_size(tmp_path):
    file = tmp_path / "a.xlsm"
    file.write_bytes(b"abc")
    got = D.fingerprint(file)
    assert got == D.Fingerprint(md5="900150983cd24fb0d6963f7d28e17f72", size=3,
                                sha256="ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    assert got.md5 == hashlib.md5(b"abc").hexdigest()


def stream(path: Path, text: str):
    Path(f"{path}:Zone.Identifier").write_text(text, encoding="ascii")


def test_zone_id_is_none_without_the_stream(tmp_path):
    file = tmp_path / "a.xlsm"
    file.write_bytes(b"x")
    assert D.zone_id(file) is None


def test_zone_id_reads_the_zone_of_the_stream(tmp_path):
    file = tmp_path / "a.xlsm"
    file.write_bytes(b"x")
    stream(file, "[ZoneTransfer]\r\nZoneId=3\r\nReferrerUrl=https://example.invalid/x\r\n")
    assert D.zone_id(file) == 3


def test_zone_id_is_none_when_the_stream_is_not_understood(tmp_path):
    file = tmp_path / "a.xlsm"
    file.write_bytes(b"x")
    stream(file, "garbage\r\nZoneId=abc\r\n")
    assert D.zone_id(file) is None


def test_zone_names():
    assert [D.zone_name(n) for n in (0, 1, 2, 3, 4, 9)] == ["local computer", "intranet", "trusted sites", "Internet",
                                                           "restricted sites", "unknown"]


@pytest.mark.parametrize("file, locations, expected", [
    ("C:\\Work\\Book\\a.xlsm", [("C:\\Work\\Book", False)], True),
    ("c:\\work\\book\\a.xlsm", [("C:\\Work\\Book\\", False)], True),                 # case and trailing separator
    ("C:\\Work\\Book\\sub\\a.xlsm", [("C:\\Work\\Book", True)], True),               # subfolders allowed
    ("C:\\Work\\Book\\sub\\a.xlsm", [("C:\\Work\\Book", False)], False),             # subfolders not allowed
    ("C:\\Other\\a.xlsm", [("C:\\Work\\Book", True)], False),
    ("C:\\Work\\Bookshelf\\a.xlsm", [("C:\\Work\\Book", True)], False),              # a name that only starts the same way
    ("C:\\Work\\Book\\a.xlsm", [], False),
])
def test_in_trusted_location(file, locations, expected):
    assert D.in_trusted_location(file, locations) is expected


def test_mask_home_ignores_case_and_only_masks_the_home():
    text = "C:\\USERS\\Someone\\Downloads\\a.xlsm and c:\\users\\someone\\x and C:\\Users\\SomeoneElse\\y"
    masked = D.mask_home(text, "C:\\Users\\someone")
    assert masked == "%USERPROFILE%\\Downloads\\a.xlsm and %USERPROFILE%\\x and C:\\Users\\SomeoneElse\\y"


def test_mask_home_with_no_home_changes_nothing():
    assert D.mask_home("C:\\x", "") == "C:\\x"


def test_tail_gives_the_last_lines(tmp_path):
    file = tmp_path / "log"
    file.write_text("\n".join(f"line {n}" for n in range(1, 8)) + "\n", encoding="ascii")
    assert D.tail(file, 3) == ["line 5", "line 6", "line 7"]
    assert D.tail(file, 50)[0] == "line 1"


def test_tail_leaves_out_the_inputs_of_a_run_so_its_run_line_stays(tmp_path):
    file = tmp_path / "jira2dashboard.log"
    params = "".join(f"2026-11-11 19:27:09.078 PARAM P{n} = (set)\n" for n in range(35))
    file.write_text("2026-11-11 19:27:09.062 RUN button Extract from Jira and compute | Excel 16.0\n" + params +
                    "2026-11-11 19:27:09.117 ERROR J2D-CFG-002 | Btn | vbObjectError+608 | x\n"
                    "2026-11-11 19:27:09.125 RUN end stopped J2D-CFG-002 | 0.133 s\n", encoding="ascii")
    assert [line.split(" ", 3)[2] for line in D.tail(file, 10)] == ["RUN", "ERROR", "RUN"]


def test_tail_of_an_empty_or_missing_file_is_empty(tmp_path):
    empty = tmp_path / "empty"
    empty.write_text("", encoding="ascii")
    assert D.tail(empty, 5) == [] and D.tail(tmp_path / "missing", 5) == []


def test_the_stable_name_is_the_delivered_workbook_name():
    assert D.DELIVERED_NAME == "jira2dashboard.xlsm"


@pytest.mark.parametrize("name, stable", [("jira2dashboard.xlsm", True), ("Jira2Dashboard.XLSM", True),
                                          ("jira2dashboard (1).xlsm", False), ("copy of jira2dashboard.xlsm", False)])
def test_name_line_says_whether_the_name_is_the_stable_one(name, stable):
    line = D.name_line(name)
    assert line.startswith(f"Name: {name} ")
    assert ("(the stable name)" in line) is stable
    assert ("not the stable name jira2dashboard.xlsm" in line and "would miss it" in line) is (not stable)


def test_vba_log_path_follows_the_rule_of_modlog():
    assert D.vba_log_path(Path("C:\\D\\a.xlsm"), "C:\\L", "C:\\T") == Path("C:\\D\\logs\\jira2dashboard.log")
    assert D.vba_log_path(None, "C:\\L", "C:\\T") == Path("C:\\L\\jira2dashboard\\logs\\jira2dashboard.log")
    assert D.vba_log_path(None, "", "C:\\T") == Path("C:\\T\\jira2dashboard\\logs\\jira2dashboard.log")


def readings(**changes):
    base = dict(office={}, policies={}, trusted_locations=[], security_products=None, interpreters=[])
    return D.Readings(**{**base, **changes})


def test_security_products_not_available_or_none_registered_are_said_apart(tmp_path):
    kwargs = dict(home="", local_app_data=str(tmp_path), temp=str(tmp_path))
    assert "Security Center not available" in D.report(None, readings(security_products=None), **kwargs)
    assert "none registered" in D.report(None, readings(security_products=[]), **kwargs)
    assert "- Some Product" in D.report(None, readings(security_products=["Some Product"]), **kwargs)


def test_every_section_is_there_even_when_nothing_was_read(tmp_path):
    text = D.report(None, readings(), home="", local_app_data=str(tmp_path), temp=str(tmp_path))
    for title in ("[File]", "[Origin]", "[Office]", "[Security products]", "[Interpreters of the fake Jira]", "[VBA log"):
        assert title in text


def test_the_python_log_is_never_read(tmp_path):
    workbook = tmp_path / "jira2dashboard.xlsm"
    folder = tmp_path / "logs"
    folder.mkdir(parents=True)
    (folder / "python.log").write_text("SECRET-IN-PYTHON-LOG\n", encoding="ascii")
    (folder / "jira2dashboard.log").write_text("2026-11-06 10:00:00.000 RUN a line of the VBA log\n", encoding="ascii")
    text = D.report(workbook, readings(), home="", local_app_data=str(tmp_path), temp=str(tmp_path))
    assert "a line of the VBA log" in text and "SECRET-IN-PYTHON-LOG" not in text


def stream_zone(path: Path, zone: int):
    Path(f"{path}:Zone.Identifier").write_text(f"[ZoneTransfer]\r\nZoneId={zone}\r\n", encoding="ascii")


def advice(tmp_path, ran, *, zone=None, **policies):
    file = tmp_path / "jira2dashboard.xlsm"
    file.write_bytes(b"x")
    if zone is not None:
        stream_zone(file, zone)
    return " ".join(D.macro_advice(ran, file, readings(policies=policies)))


def test_macros_that_do_not_run_even_with_office_relaxed_point_at_the_endpoint_protection(tmp_path):
    text = advice(tmp_path, False)
    assert "even with Office's security relaxed" in text and "endpoint protection" in text
    assert "run.ps1 diagnose" in text and "Audit" in text


def test_macros_that_run_with_the_mark_of_the_web_point_at_unblocking(tmp_path):
    text = advice(tmp_path, True, zone=3)
    assert "run when Office's security is relaxed" in text
    assert "Mark of the Web" in text and "Unblock" in text and "Trusted Location" in text


@pytest.mark.parametrize("policy, value, word", [("BlockContentExecutionFromInternet (policy)", "1", "policy blocks macros"),
                                                 ("VBAWarnings (policy)", "4", "VBAWarnings=4"),
                                                 ("VBAWarnings (user setting)", "3", "VBAWarnings=3")])
def test_macros_that_run_with_a_restrictive_office_setting_name_it(tmp_path, policy, value, word):
    assert word in advice(tmp_path, True, **{policy: value})


def test_permissive_office_settings_are_not_blamed(tmp_path):
    text = advice(tmp_path, True, **{"VBAWarnings (user setting)": "1", "VBAWarnings (policy)": "not set",
                                      "BlockContentExecutionFromInternet (policy)": "not set"})
    assert "Mark of the Web" not in text and "VBAWarnings=" not in text and "policy blocks" not in text
    assert "Nothing in Office explains a block" in text and "endpoint protection" in text


# --- the findings, with their support codes -----------------------------------------------------------------------------------

def codes(tmp_path, workbook=None, **kw):
    return [code for code, _ in D.findings(workbook, kw.pop("readings", readings()), str(tmp_path), str(tmp_path))]


def a_workbook(tmp_path, name="jira2dashboard.xlsm", zone=None):
    file = tmp_path / name
    file.write_bytes(b"x")
    if zone is not None:
        stream_zone(file, zone)
    return file


def test_no_workbook_is_a_finding_with_its_code(tmp_path):
    assert "J2D-SYS-020" in codes(tmp_path)
    assert "J2D-SYS-020" not in codes(tmp_path, a_workbook(tmp_path))


def test_the_mark_of_the_web_is_a_finding_with_its_code(tmp_path):
    found = codes(tmp_path, a_workbook(tmp_path, zone=3))
    assert "J2D-SYS-021" in found and "J2D-SYS-028" not in found                   # a cause was found in the file
    assert "J2D-SYS-021" not in codes(tmp_path, a_workbook(tmp_path))


def test_a_renamed_copy_and_a_folder_that_is_not_trusted_are_hints(tmp_path):
    found = codes(tmp_path, a_workbook(tmp_path, name="jira2dashboard (1).xlsm"))
    assert "J2D-SYS-027" in found and "J2D-SYS-022" in found
    assert "J2D-SYS-027" not in codes(tmp_path, a_workbook(tmp_path))


@pytest.mark.parametrize("policy, value, code", [("BlockContentExecutionFromInternet (policy)", "1", "J2D-SYS-023"),
                                                 ("VBAWarnings (policy)", "4", "J2D-SYS-024"),
                                                 ("VBAWarnings (user setting)", "3", "J2D-SYS-024")])
def test_a_restrictive_office_setting_is_a_finding_with_its_code(tmp_path, policy, value, code):
    found = codes(tmp_path, a_workbook(tmp_path), readings=readings(policies={policy: value}))
    assert code in found and "J2D-SYS-028" not in found


def test_nothing_in_office_points_at_the_endpoint_protection(tmp_path):
    assert "J2D-SYS-028" in codes(tmp_path, a_workbook(tmp_path))


def test_only_a_product_other_than_defender_is_a_finding(tmp_path):
    assert "J2D-SYS-025" not in codes(tmp_path, a_workbook(tmp_path), readings=readings(security_products=["Windows Defender"]))
    assert "J2D-SYS-025" not in codes(tmp_path, a_workbook(tmp_path), readings=readings(security_products=None))
    assert "J2D-SYS-025" in codes(tmp_path, a_workbook(tmp_path), readings=readings(security_products=["Windows Defender", "Some EDR"]))


def test_no_vba_log_is_a_finding_until_there_is_one(tmp_path):
    workbook = a_workbook(tmp_path)
    assert "J2D-SYS-026" in codes(tmp_path, workbook)
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "jira2dashboard.log").write_text("x\n", encoding="ascii")
    assert "J2D-SYS-026" not in codes(tmp_path, workbook)


def test_the_report_opens_with_the_findings_and_lists_the_codes_to_quote(tmp_path):
    text = D.report(a_workbook(tmp_path, zone=3), readings(), home="", local_app_data=str(tmp_path), temp=str(tmp_path))
    assert text.index("[Findings]") < text.index("[File]")
    assert "J2D-SYS-021: Mark of the Web" in text
    assert re.search(r"Codes to quote to the support: J2D-SYS-021.*J2D-SYS-022", text)


def test_the_advice_of_the_macro_check_carries_the_codes(tmp_path):
    assert "J2D-SYS-029" in advice(tmp_path, False)
    assert "J2D-SYS-021" in advice(tmp_path, True, zone=3)
    assert "J2D-SYS-023" in advice(tmp_path, True, **{"BlockContentExecutionFromInternet (policy)": "1"})
    assert "J2D-SYS-024" in advice(tmp_path, True, **{"VBAWarnings (user setting)": "3"})
    assert "J2D-SYS-028" in advice(tmp_path, True)
