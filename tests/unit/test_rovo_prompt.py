"""The text behind the button Prompt for Rovo (rovo_prompt.py): what Rovo is asked, and the CSV it must answer."""
from jira2dashboard import rovo_prompt as RP
from jira2dashboard.params.model import SPECS


def test_every_rule_talks_about_a_parameter_sheet_of_the_workbook():
    names = {spec.name for spec in SPECS}
    assert set(RP.SHEET_RULES) <= names
    assert "Config" not in RP.SHEET_RULES                      # Config has its own part: one line per parameter, read live


def test_the_spec_has_one_config_part_and_one_sheet_part_per_rule_in_order():
    rows = RP.spec_rows()
    assert [name for part, name, _text in rows if part == RP.SHEET] == list(RP.SHEET_RULES)
    assert [part for part, _name, _text in rows].count(RP.CONFIG) == 1
    parts = [part for part, _name, _text in rows]
    assert parts.index(RP.CONFIG) < parts.index(RP.SHEET)             # Config first, then the sheets, then the format
    assert {part for part, _name, _text in rows} == {RP.TEXT, RP.CONFIG, RP.SHEET, RP.KEEP}
    assert all(part and (part != RP.SHEET or name) for part, name, _text in rows)        # a row is never read without a Part


def test_rovo_is_asked_for_a_csv_with_the_columns_the_workbook_can_read_back():
    text = [text for part, _name, text in RP.spec_rows() if part == RP.TEXT]
    assert ",".join(RP.CSV_COLUMNS) in text and RP.CSV_COLUMNS == ("Sheet", "Row", "Column", "Value", "Note")
    joined = "\n".join(text)
    assert "CSV" in joined and RP.UNKNOWN in joined
    for sheet in ("Config", *RP.SHEET_RULES):
        assert sheet in joined.split("- Sheet:")[1].split("\n")[0], sheet           # each sheet is named in the format


def test_the_connection_parameters_are_kept_out_of_the_load_and_exist_in_config():
    from jira2dashboard.params.model import CONFIG_PARAMETERS
    assert set(RP.PROTECTED) <= set(CONFIG_PARAMETERS)                             # each one is a parameter of Config
    names = [name for part, name, _text in RP.spec_rows() if part == RP.KEEP]
    assert names == list(RP.PROTECTED)
    assert {"BaseUrl", "Email", "ApiToken", "ProxyServer"} <= set(names)       # where the credentials go


def test_the_prompt_never_asks_for_or_carries_a_secret():
    text = "\n".join(text for _part, _name, text in RP.spec_rows())
    assert "TO BE TYPED BY THE USER" in text and "Never write an API token" in text
    assert "http" not in text.lower() and "@" not in text         # no address and no e-mail: only names and meanings
