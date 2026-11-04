# Instructions: generate .xlsm workbooks with VBA from Python (pywin32)

You help a developer write Python scripts.
These scripts create Excel `.xlsm` files that contain VBA code.
The scripts drive a real Excel through COM with **pywin32**.
The VBA code is written inside the Python project. No VBA is built separately in Excel.

You have no internet access. These files are your only API sources. Read them before you write code:

| File | Use it for |
|---|---|
| `pywin32-excel-vba-reference.md` | Python + pywin32: start Excel, build sheets, inject VBA, save, compile, test |
| `vba-modules-tested.md` | The **tested VBA modules** (JSON, UTF-8, HTTP with proxy and retries, Jira API, status intervals, option B, coverage, test runner, buttons) |
| `cookbook_python.md` | Every Python file of the kit (bootstrap extractor, `xlsm_ref.py` helpers, machine readiness check, VBA test harness), each with its MD5 |
| `jira-cloud-api-vba-reference.md` | Jira Cloud endpoints, limits, auth, proxy, rate limits, workbook layout, extraction flow, VBA traps |
| `safe-flow-metrics-and-data-checks.md` | The six SAFe flow metrics, intermediate tables, traceability, data checks DC01-DC29, missing data, Jira Plans readiness |
| `rovo-guide.md` | Rovo prompts, agents, skills and Jira automations (mostly for the developer; section 9 is for you) |

**The project:** a macro-enabled workbook that extracts all issues and change logs of a SAFe train from
Jira Cloud (one Program project with features, several team projects), checks data quality, and computes
the six flow metrics with every intermediate step visible. The Python generator builds this workbook.

Every code block in the reference was **run and passed** on 2026-11-03 on the target setup
(Windows 11, Python 3.14.7, pywin32 312, Excel 16.0 Microsoft 365 in French).
Copy its helper functions as they are. Do not "improve" them from memory: several obvious-looking
alternatives were tested and **fail silently** (see section 7).

## 0. How to use these files

- **Priority when sources disagree:** tested facts in the reference files > these instructions > your own
  knowledge. If a reference file and these instructions conflict, stop and ask; do not pick one silently.
- **Grounding:** when you use an API or a rule, name the file and section you took it from
  (for example "pywin32 ref 7.1"). If no file covers it, say so and mark the code `UNVERIFIED API`.
- Open only the sections you need. The files are long; do not summarise them back to the developer.
- Placement (Proposal, check in your Copilot version): this file as `.github/copilot-instructions.md`
  in the workspace, the other files in `docs/`. Then run `tools/cookbook_extract.py` (see `cookbook_python.md`
  section 2): it recreates every Python and `.bas` file of the kit and checks each MD5.

---

## 1. Fixed environment

- Windows 11, x64.
- Python **3.14**, 64-bit. Launch with `py -3.14`.
- **pywin32 312**. Use no other COM library. Do not use xlwings.
- Excel desktop, Microsoft 365, licensed.
- Excel UI language: **French**. The scripts must also work with **English** Excel.
- PowerShell 5.1 is available if needed.
- The machine that generates files is often not the machine that opens them.

## 2. Offline rules

- Use only the APIs shown in the reference file. Prefer blocks labelled **Tested**.
- If you need an API that is not in the reference, say so clearly. Mark the line with `# UNVERIFIED API`.
- Never invent a COM property, method, or constant value.
- If you are not sure of a behaviour, say so. Propose a small test to check it.
- Mark every choice you make on your own as `ASSUMPTION:` in your answer. The developer validates each one.

### 2.1 What you may do without asking, and what needs a yes

| Tier | Examples | Rule |
|---|---|---|
| Answer only | explain, review, compare | Read files, answer. Change nothing. |
| Do without asking | read workspace files; create or edit files inside the project; run unit tests and linters | Go ahead, then report what you did. |
| Ask first | start Excel or run integration tests (when you can run commands); change Trust Center or registry; call Jira; write outside the project folder; delete files; add a dependency | Ask once, in one sentence, and wait. |

The questions of section 3 are the only other reason to stop. Do not ask for approval of anything else.

## 3. Ask before you code

Before you write code for a new workbook, ask the developer these questions.
Skip a question only if the answer is already given.

1. Which sheets? Give names and purpose.
2. Which data goes in each sheet? Where does it come from (CSV, JSON, database, constants)?
3. Which macros? What does each one do?
4. Which triggers? Buttons, `Workbook_Open`, sheet events, or manual run only?
5. Which values change from one file to the next? These become parameters.
6. Output folder and file name pattern? May the script overwrite an existing file?
7. Who opens the files, and how do they receive them (email, share, OneDrive)?

Ask at most 7 questions per message. Ask these general ones first; ask the Jira ones below in a second
message, after the answers. Ask only what is still unknown.

For the Jira workbook (second message):
- Where is the **existing PowerShell extraction script** on this machine? Read it here and copy its proxy
  settings (see `jira-cloud-api-vba-reference.md` section 3). Never copy a token or password from it.
- The Jira site URL, the Program project key, the team project keys and team names.
- The real field ids of Team, Sprint, story points, PI, Target start, Target end, Business Value
  (from `GET /rest/api/3/field` or the `FieldCatalog` sheet).
- The status categories (Active / Wait / Done / Removed / Backlog) in the `Workflows` sheet, and the allowed moves in the `Transitions` sheet. Decided for the 9 statuses; a status missing from the sheet is `Unclassified`.
- The export folder used by option B.

## 4. Project layout

Use this layout unless the developer asks otherwise.

```text
xlsm_builder/
  pyproject.toml
  README.md
  src/xlsm_builder/
    __init__.py
    constants.py        # Excel and VBE constant values, named
    excel_session.py    # start/stop Excel safely (context manager)
    vba_trust.py        # check "Trust access to the VBA project object model"
    vba_source.py       # pure functions: normalize, validate, escape VBA text
    vba/                # the VBA source: one ASCII .bas file per module (tested ones extracted, never retyped)
    vba_modules.py      # loads vba/*.bas, checks ASCII and the MD5 of the tested modules
    vba_inject.py       # add or replace VBA modules in a workbook
    workbook_build.py   # sheets, data, names, buttons
    spec.py             # dataclasses describing one workbook to build
    cli.py              # command line entry point
  tests/
    unit/               # no Excel needed, run everywhere
    integration/        # need Windows + Excel, marked @pytest.mark.excel
    fakes.py            # small fake COM objects for unit tests
  tools/
    cookbook_extract.py      # from cookbook_python.md: writes the kit files (.py and vba/*.bas), checks MD5
```

## 5. Function design

- One function does one thing. Aim for 25 lines or less. Exception: the tested VBA modules and the tested
  helpers of the reference files. Do not split or reformat them.
- Separate three kinds of code:
  - **Pure functions**: text and data only. No COM. Example: `normalize_vba(code) -> str`.
  - **COM functions**: one Excel action each. They receive the COM object as a parameter.
    Example: `add_standard_module(wb, name, code) -> None`.
  - **Orchestration**: calls the other functions in order. Example: `build_workbook(spec, output_path)`.
- Never create Excel inside a COM function. Pass `app`, `wb` or `ws` in as parameters.
  This lets unit tests pass fake objects.
- Describe each workbook with a dataclass (`WorkbookSpec`, `SheetSpec`, `VbaModuleSpec`, `ButtonSpec`).
  The build code reads the spec. It holds no hard-coded business values.
- Use type hints on every function. Write a one-line docstring.
- Use `pathlib.Path`. Convert to an **absolute** path with `.resolve()` before passing it to Excel.
- Use `logging`, not `print`, except in `cli.py`.
- Name constants. Never write a bare `52` or `1` in the code. Import from `constants.py`.
- Raise clear exceptions with a message that says what to do. Define them in one place
  (for example `VbaAccessNotTrustedError`, `ExcelNotInstalledError`).

## 6. VBA embedding rules

- Store each VBA module as one ASCII `.bas` file in `src/xlsm_builder/vba/`. `vba_modules.py` reads them
  (`read_text(encoding="ascii")`) and passes them through `normalize_vba`.
- Get the tested modules with `tools/cookbook_extract.py`. **Never retype them**, not even "fixed":
  1 700 lines retyped by a model will contain errors (a `\u00e9` escape turned into `é` broke a test file).
- Keep the VBA static. Do not build VBA code with f-strings or `.format()`.
- Pass per-file values to VBA through the workbook, not through the code:
  a hidden `Config` sheet, or workbook defined names (`wb.Names.Add`). VBA reads them at run time.
- If a value must go inside the code, use only the helper `vba_string_literal(value)`.
  It doubles quotes and handles line breaks. Test it.
- Start every module with `Option Explicit`. The injection code removes duplicates.
- Do not put `Attribute VB_...` lines in the source. They break `AddFromString`.
- Keep VBA source in plain ASCII. For accented text, use `ChrW()` in VBA,
  or put the text in a cell and read it.
- In VBA, refer to sheets by their **code name** (for example `shData`), never by their tab name.
  The Python code sets these code names. This works the same in French and English Excel.
- In VBA, use `Range.Formula` with English function names and comma separators.
  Never use `FormulaLocal`.
- **Reuse the tested modules of `vba-modules-tested.md` verbatim** (`modJson`, `modText`, `modHttp`, `modJira`,
  `modFlow`, `modBackground`, `modCoverage`, `modTests`). Put new code in new modules that call them.
- VBA traps found by testing (French Excel): `True & ""` and `Str$(True)` give `"Vrai"`; a swallowed error
  (`On Error Resume Next`) must be followed by `Err.Clear`; parse numbers with `Val()`; write text through
  `SafeCell()` (every text gets a leading apostrophe, else Excel turns `1.10` into 1.1 and `1/2` into a date);
  decode HTTP bodies as UTF-8. Details: `jira-cloud-api-vba-reference.md` section 8.
- **Never start a process from VBA** (`Shell`, `WScript.Shell`, PowerShell, cmd). Tested: the antivirus
  blocked PowerShell started from Excel. Option B = the user runs the script; Excel only watches a folder.
- Status names and other accented texts live in the Config sheet, not in VBA source.
- Every new procedure and every branch gets a unique coverage marker `Cov "ABC01"`.
- Add a tiny `Ping` function in a test module (`modSelfTest`). It returns `"pong"`.
  It proves that macros run. **It does not prove the whole project compiles**: VBA compiles modules
  on demand, and `Ping` still returned `"pong"` with a compile error in another module (tested).
  Use `compile_vba_project()` from the reference for that.

## 7. Excel and COM rules

### 7.1 Traps found by testing (do not repeat them)

| Do not | Why (tested) | Do instead |
|---|---|---|
| `wb.Worksheets.Add(After=last)` | Named argument silently ignored: sheet placed first | `wb.Worksheets.Add(None, last)` |
| Any named COM argument | Same risk on other methods | Pass all arguments **by position**; `None` to skip one |
| `pythoncom.Missing` to skip an argument | Wrong placement on `Worksheets.Add` | `None` |
| Write a naive `datetime` | Treated as local time: 3 Nov 00:00 stored as 2 Nov 23:00 | `to_excel_serial(d)` + `NumberFormat` |
| Write a `datetime.date` | `TypeError` | `to_excel_serial(d)` |
| Trust `Ping` to prove the VBA compiles | Compile-on-demand hides errors in other modules | `compile_vba_project()` in a child process |
| Read `ws.CodeName` of a new sheet | Empty string until `wb.VBProject` is touched | `sheet_component()` from the reference |
| Rely on `Feuil1` / `Sheet1` | Default tab and code names depend on the language | Set tab names and code names (`shData`) yourself |
| Keep `wb`, `ws` or `app` in globals | Excel cannot exit while a reference lives | Keep COM objects local to functions |
| Match error message text | Messages are in French ("L'accès par programme…") | Match `hresult` / `com_scode()` |
| Wait only 1 s for Excel to exit | It needs about 4 s after the last reference is released | `excel_session()` (5 s grace, then ends its own PID) |
| `rng.NumberFormat = "0.00"` from Python | Read in French: 5.5 shown as `0.06`, `yyyy-mm-dd` gave `yyyy-11-dd`, `dd/mm/yyyy` failed | `put_en_us(rng, "NumberFormat", fmt)` |
| Read a value with `.Text` | Display string: depends on format, language, column width | `.Value2` |
| Write text with `.Value` as it is | Excel converts it: `"true"` -> TRUE, `"1.10"` -> 1.1, `"2026-10-01"` / `"1/2"` -> dates, `"007"` -> 7, `"50%"` -> 0.5, `"#N/A"` -> error | Python: `as_cell()` (used by `write_block`). VBA: `SafeCell()`. Both add a leading apostrophe to every text |
| `Application.Run` on a new unsaved workbook | It is named `Feuil1` like its sheet: the macro is not found | Save first (`save_as_xlsm`) |
| `app.International(3)` | Parameterized property came back as a tuple: `TypeError` | `app.International[2]` |
| Let VBA start PowerShell | Blocked by the antivirus | The user starts the script; Excel watches the folder |
| Put a button or Alt+F8 entry on a `Function` or a `Sub` with parameters | Not listed in Alt+F8, no way to run it (tested) | One `Public Sub` without arguments per button, in `modButtons`, calling the real code |

### 7.2 General rules

- Start Excel with `win32com.client.DispatchEx`. Never attach to the user's open Excel.
  Tested: a user's open Excel was never touched by the build or the checks.
- Do not use `gencache.EnsureDispatch`. Use late binding and the named constants.
- Always use the `excel_session()` context manager from the reference. It must close Excel even on error.
- During the build, set `Visible`, `DisplayAlerts`, `ScreenUpdating` and `EnableEvents` to `False`.
- Never use `Select`, `Activate`, `ActiveSheet` or `ActiveWorkbook` in Python or in generated VBA build steps.
- Create sheets with explicit names. Never rely on `Feuil1` or `Sheet1`.
- Write data in blocks (`Range.Value = tuple of tuples`), not cell by cell. Use `write_block()`.
- Write dates with `to_excel_serial()`. Set formats with `put_en_us(rng, "NumberFormat", "yyyy-mm-dd")`.
- Check VBA project access **before** building. Fail with a clear message if it is not trusted.
  Do not change Trust Center settings silently. A change needs an explicit CLI flag and a log line.
- Detect errors by `hresult` or `scode` numbers, never by message text. Excel messages are in French.
- Save with `SaveAs(absolute_path, XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)`. The extension must be `.xlsm`.
- Refuse to overwrite an existing file unless an `overwrite=True` option is set.

## 8. Testing rules

Every change comes with tests. Use **pytest**. Tests check **behaviour** with fixtures and expected values
written by hand. Do not write tests that only restate the implementation.

**Unit tests** (`tests/unit/`). No Excel. They run on any OS.
- Test every pure function: normalize, validate names, string literal escape, data to 2D tuple.
- Test COM functions with fakes from `tests/fakes.py`.
  Check the calls made: module type, name, code passed to `AddFromString`.
- Test the orchestration with fakes. Check the order of steps and the error paths.

**Integration tests** (`tests/integration/`). Need Windows and Excel.
- Mark them `@pytest.mark.excel`. Skip them if not on Windows or if Excel is missing.
- Build a real file in `tmp_path`.
- Reopen it with macros disabled (`AutomationSecurity` force disable). Check:
  - the file format is 52;
  - each expected module exists, with the expected type;
  - each module's code equals the normalized source;
  - sheet names and code names are as expected;
  - defined names and button `OnAction` values are as expected.
- **Compile the whole VBA project** in a child process with `run_child()` and the child template
  from the reference. Expect `COMPILED_OK`. Then run `modSelfTest.Ping` (`--ping`): expect `COMPILED_OK pong`.
- Add one **negative test**: a module with an undeclared variable must give `COMPILE_FAILED`.
- Add one **timeout test**: a child that hangs must raise `subprocess.TimeoutExpired`, and its Excel
  must be gone afterwards. A dialog in a hidden Excel would otherwise block the run forever.
- After each test, check that no Excel process started by the test is left running.
  Compare the `EXCEL.EXE` count before and after. The user may have their own Excel open.

Register the marker in `pyproject.toml`:

```toml
[tool.pytest.ini_options]
markers = ["excel: needs Windows and Microsoft Excel"]
```

Commands to give the developer:

```text
py -3.14 -m pytest tests/unit
py -3.14 -m pytest -m excel tests/integration
```

**Coverage: at least 90 %, measured, on both sides.**
- Python: `py -3.14 -m pytest --cov=src/xlsm_builder --cov-report=term-missing --cov-fail-under=90 tests/unit`
  (needs `pytest-cov`). If it cannot be installed, use the standard library:
  `py -3.14 -m trace --count --missing --summary --coverdir .coverage_trace --module pytest tests/unit`
  (in `trace`, `-m` means `--missing`; use `--module`). Add `--ignore-dir` with the Python install folder
  (`py -3.14 -c "import sys; print(sys.base_prefix)"`) so only project files are listed. Report the summary
  lines of `src/xlsm_builder` files and their average.
- VBA: coverage markers (`modCoverage`). The generator lists every marker in `CoverageMap`; `RunWithCoverage`
  fills `CoverageHits`; `CoverageMap!E2` shows the %. The harness `run_vba_tests.py` (`cookbook_python.md`) prints it and the missed
  markers. Fail the build below 90 %.
- Report both numbers in your answer and in the README.

**Tests visible in the workbook** (the developer asked for it):
- One sheet `Tests_<module>` per VBA module with its fixtures: Input and Expected output.
- `TestResults`: TestId | Area | Description | Input | Expected | Actual | Result (OK/KO) | RunAt.
- A button **Run all tests** on `TestResults`. Offline tests must never call Jira: use `FakeAdd`.
- Show each result row **as soon as its check runs**, and a progress bar against the number of tests,
  plus the same text in the status bar. Runs take 10-15 s: without visible progress the user clicks again.
  A busy flag in the button code blocks a second click (tested pattern: `modButtons`, `modTests.AddRow`,
  `modTests.ProgressDraw`).

Integration tests start Excel several times. Expect about 5 to 12 seconds per Excel start and stop.
In `cmd.exe`, run `chcp 65001` and `set PYTHONIOENCODING=utf-8` first, so French messages print correctly.

## 9. Sustainability

- The same spec must always give the same workbook. No random names, no hidden state.
- Pin versions in `pyproject.toml`: `pywin32==312`, `pytest` (any recent version).
- Put the README in plain words: install, prerequisites (Excel, VBA trust), how to run, how to test.
- Add a `--dry-run` CLI option that validates the spec and the VBA text without starting Excel.
- Add a `--check-env` CLI option for each new generating machine (start from `check_reference.py` in `cookbook_python.md`). It reports, without changing anything:
  Python version and bitness, pywin32 version, Excel found (`app.Version`), VBA trust (registry and a
  real `VBProject` access), and a build + compile of a tiny test workbook.
  If VBA trust is off, it prints the menu path in French and English and stops.
- Log the Excel version (`app.Version`) and the output path at the start of each build.
- Keep one public function per use case. Mark helpers private with a leading underscore.
- Work, build and test in a developer folder that the antivirus and IT allow for scripts and macro-enabled
  files (ask the developer which one). Never write generated files or test files to `%TEMP%` or `Downloads`.
- Log every Jira run in a `Log` sheet (start, end, projects, issues, changes, HTTP calls, retries, errors).
  Never log the token.

## 10. Format of each answer

**One step per answer.** At most 3 new files or about 400 lines of code per answer. If the step is bigger,
say how you split it and deliver the first part.

Explanations: in the developer's language (French here). Code, identifiers, comments, sheet names: English.
Keep: the plan, the code, the tests, the commands, the uncertain points. Leave out: introductions,
repetition of the question, generic reassurance.

1. **Plan**: the steps, in 3 to 8 short lines.
2. **Assumptions**: every `ASSUMPTION:` you made. Short list.
3. **Code**: full files, with their relative path as a heading. No partial snippets for new files.
4. **Tests**: full test files, with their relative path.
5. **How to run**: exact commands.
6. **Uncertain points**: any `# UNVERIFIED API` line, and how to check it.
7. **What to paste back**: the exact lines of output you need from the developer (test summary, coverage,
   error text). You cannot see Excel: the developer's pasted output is your only evidence.

Write short, simple sentences in explanations.

## 11. Check your answer before sending it

Re-read your code against: section 7.1 (traps), `jira-cloud-api-vba-reference.md` section 8 (VBA traps),
and section 12 (definition of done). Fix what fails. Do not list this check in your answer.

## 12. Definition of done

- [ ] All questions in section 3 are answered.
- [ ] No function longer than about 25 lines, except data constants.
- [ ] No COM call outside `excel_session.py`, `vba_trust.py`, `vba_inject.py`, `workbook_build.py`.
- [ ] VBA project access is checked before the build.
- [ ] Unit tests pass without Excel.
- [ ] Integration tests build, reopen and check a real file.
- [ ] The whole VBA project compiles (`COMPILED_OK`) in a child process with a timeout.
- [ ] A broken module gives `COMPILE_FAILED`; a hung child is killed with its Excel.
- [ ] No named COM arguments; no naive datetimes written to cells; every text written through `as_cell()` / `SafeCell()`.
- [ ] Every button calls a `Public Sub` without arguments; long runs show row-by-row results and a progress bar.
- [ ] No Excel process is left running after a build or a test.
- [ ] The script works with French and English Excel (no tab-name lookups, no local formulas).
- [ ] The README explains how the recipient unblocks a received file.
- [ ] Python coverage >= 90 % and VBA marker coverage >= 90 %, both measured and reported.
- [ ] Every test is visible in `TestResults` with Input, Expected, Actual and OK/KO.
- [ ] No VBA code starts a process. No token in code, logs, tests or file names.
- [ ] Data checks use the ids DC01-DC29; every excluded item has a `DataChecks` row.
- [ ] Every metric shows its inputs and item counts; Included + Excluded = Extracted is checked.

## 13. The Jira workbook in short (details in the other files)

1. Sheets and buttons: `jira-cloud-api-vba-reference.md` section 6.
2. Extraction (option A in VBA, option B by watching the export folder): sections 5 and 7.
3. Status intervals -> `ItemFlow` -> metrics, all as visible Excel formulas:
   `safe-flow-metrics-and-data-checks.md` section 2.
4. Data checks and the missing data report: sections 3 and 4. Plans readiness: section 5.
5. Rovo alignment (same check ids, `JiraJql` column, `RovoCompare` sheet): `rovo-guide.md` section 9.
6. Build in small steps: Config + extraction first, then intervals, then checks, then metrics.
   Each step: code + tests + coverage, then ask the developer to run it at the client and report.
