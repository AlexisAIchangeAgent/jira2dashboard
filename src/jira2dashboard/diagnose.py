"""What `scripts\\run.ps1 diagnose` reports, as pure functions: nothing is read from Windows here (cli/diagnose.py does that).

The report is meant for a ticket to the team of the endpoint protection that blocks the macros: the hashes of the
workbook (an allow-list often asks for the MD5), where it comes from (Mark of the Web, trusted location), the Office and
its macro policies, the security products, the interpreters of the fake Jira and the end of the VBA log.
It never opens the workbook (no macro, no Config sheet), never reads python.log (it holds the JQL and Jira's messages)
and never reads the Windows event log. The home folder of the user is masked.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import ntpath
import re
from dataclasses import dataclass
from pathlib import Path

from jira2dashboard.paths import DELIVERED_NAME

ZONES = {0: "local computer", 1: "intranet", 2: "trusted sites", 3: "Internet", 4: "restricted sites"}


@dataclass(frozen=True)
class Fingerprint:
    md5: str
    sha256: str
    size: int


@dataclass(frozen=True)
class Readings:
    """What cli/diagnose.py read from Windows."""
    office: dict[str, str]
    policies: dict[str, str]
    trusted_locations: list[tuple[str, bool]]          # (folder, subfolders allowed)
    security_products: list[str] | None                # None: Security Center not available
    interpreters: list[Path]


def fingerprint(path: Path) -> Fingerprint:
    md5, sha256, size = hashlib.md5(), hashlib.sha256(), 0
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            md5.update(block)
            sha256.update(block)
            size += len(block)
    return Fingerprint(md5.hexdigest(), sha256.hexdigest(), size)


def zone_id(path: Path) -> int | None:
    """The ZoneId of the Mark of the Web (NTFS stream Zone.Identifier); None when the stream is absent or not understood."""
    try:
        text = Path(f"{path}:Zone.Identifier").read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    for line in text.splitlines():
        key, _, value = line.partition("=")
        if key.strip() == "ZoneId" and value.strip().isdigit():
            return int(value.strip())
    return None


def zone_name(zone: int) -> str:
    return ZONES.get(zone, "unknown")


def in_trusted_location(file: str | Path, locations: list[tuple[str, bool]]) -> bool:
    """Is the file in one of the Trusted Locations of Excel (its folder, or below it when subfolders are allowed)?"""
    def norm(text: str) -> str:
        return ntpath.normcase(ntpath.normpath(text))
    folder = norm(ntpath.dirname(str(file)))
    for location, subfolders in locations:
        base = norm(location)
        if folder == base or (subfolders and folder.startswith(base + "\\")):
            return True
    return False


def mask_home(text: str, home: str) -> str:
    """The home folder of the user becomes %USERPROFILE% (a folder with the same start, like name2, is left alone)."""
    home = home.rstrip("\\/")
    if not home:
        return text
    return re.sub(re.escape(home) + r"(?![A-Za-z0-9_.-])", lambda _: "%USERPROFILE%", text, flags=re.IGNORECASE)


def tail(path: Path, count: int) -> list[str]:
    """The last `count` lines of the log, its PARAM lines left out: about 35 at the start of each run (the inputs, masked),
    they would push the line RUN, which names the button, out of a short tail."""
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []
    lines = [line for line in lines if line.split(" ", 3)[2:3] != ["PARAM"]]
    return lines[-count:] if count > 0 else []


def vba_log_path(workbook: Path | None, local_app_data: str, temp: str) -> Path:
    """Where modLog writes: logs\\ next to the workbook; without a workbook, %LOCALAPPDATA%\\jira2dashboard\\logs (the
    fallback of modLog), or the TEMP folder when LOCALAPPDATA is empty."""
    if workbook is not None:
        return workbook.parent / "logs" / "jira2dashboard.log"
    return Path(local_app_data or temp) / "jira2dashboard" / "logs" / "jira2dashboard.log"


def name_line(name: str) -> str:
    """The file name against the stable one: an allow-list on the name or the path misses a renamed copy (a second download)."""
    if name.lower() == DELIVERED_NAME:
        return f"Name: {name} (the stable name)"
    return f"Name: {name} (not the stable name {DELIVERED_NAME}: an allow-list on the name or the path would miss it)"


VBA_USER = "VBAWarnings (user setting)"
VBA_POLICY = "VBAWarnings (policy)"
BLOCK_INTERNET = "BlockContentExecutionFromInternet (policy)"
RESTRICTIVE_VBA_WARNINGS = {"3", "4"}                  # 3: only macros signed by a trusted publisher, 4: none


NO_WORKBOOK, MOTW, NOT_TRUSTED, BLOCKED_BY_POLICY, UNSIGNED_DISABLED, OTHER_PRODUCT, NO_VBA_LOG, NOT_STABLE_NAME, \
    NOTHING_IN_OFFICE, MACROS_DID_NOT_RUN = (f"J2D-SYS-0{n}" for n in (20, 21, 22, 23, 24, 25, 26, 27, 28, 29))   # error_codes.TABLE
DEFENDER = re.compile(r"(windows|microsoft) defender", re.IGNORECASE)


def macro_advice(ran: bool, workbook: Path | None, readings: Readings) -> list[str]:
    """What the macro check of check-machine means. `ran`: a macro of the workbook ran with Office's own security relaxed
    for the check (AutomationSecurity low). If it did not, Office is not the cause: something outside it stops the macros."""
    if not ran:
        return [f"{MACROS_DID_NOT_RUN}: The macros did not run even with Office's security relaxed: something outside Office "
                "stops them, most likely the endpoint protection.",
                "Next: run scripts\\run.ps1 diagnose for the ticket, and ask the security team to put a test machine in Audit "
                "mode to read the alert (class, process, rule)."]
    reasons = []
    zone = zone_id(workbook) if workbook is not None and workbook.is_file() else None
    if zone is not None and zone >= 3:
        reasons.append(f"{MOTW}: the file has the Mark of the Web: right-click > Properties > Unblock, or put it in a Trusted Location.")
    if readings.policies.get(BLOCK_INTERNET) == "1":
        reasons.append(f"{BLOCKED_BY_POLICY}: a policy blocks macros of files from the Internet: unblock the file or use a Trusted Location.")
    for label in (VBA_POLICY, VBA_USER):
        value = readings.policies.get(label)
        if value in RESTRICTIVE_VBA_WARNINGS:
            reasons.append(f"{UNSIGNED_DISABLED}: the macro setting of Office disables unsigned macros (VBAWarnings={value}): use a Trusted Location.")
    if not reasons:
        reasons.append(f"{NOTHING_IN_OFFICE}: Nothing in Office explains a block: if the macros still do not run by hand, the endpoint "
                       "protection is the likely cause (run.ps1 diagnose, ask the security team for its Audit mode).")
    return ["The macros of this file run when Office's security is relaxed for this check. If they do not run by hand, "
            "look at Office:", *[f"- {reason}" for reason in reasons]]


def findings(workbook: Path | None, readings: Readings, local_app_data: str, temp: str) -> list[tuple[str, str]]:
    """What the readings say, as (support code, text): the codes the person quotes to the support (error_codes.TABLE).
    A cause of a block (Mark of the Web, a policy, a macro setting) or a hint (name, trusted location, a security product)."""
    found: list[tuple[str, str]] = []
    cause = False
    if workbook is None or not workbook.is_file():
        found.append((NO_WORKBOOK, "No workbook found: build one or give --workbook FILE."))
    else:
        if workbook.name.lower() != DELIVERED_NAME:
            found.append((NOT_STABLE_NAME, f"The file is named {workbook.name}, not {DELIVERED_NAME}."))
        zone = zone_id(workbook)
        if zone is not None and zone >= 3:
            cause = True
            found.append((MOTW, f"Mark of the Web (ZoneId={zone}, {zone_name(zone)})."))
        if not in_trusted_location(workbook, readings.trusted_locations):
            found.append((NOT_TRUSTED, f"Not in a Trusted Location of Excel ({len(readings.trusted_locations)} read)."))
    if readings.policies.get(BLOCK_INTERNET) == "1":
        cause = True
        found.append((BLOCKED_BY_POLICY, "A policy blocks the macros of files from the Internet."))
    for label in (VBA_POLICY, VBA_USER):
        value = readings.policies.get(label)
        if value in RESTRICTIVE_VBA_WARNINGS:
            cause = True
            found.append((UNSIGNED_DISABLED, f"{label} = {value}: unsigned macros are disabled."))
    others = [name for name in readings.security_products or [] if not DEFENDER.search(name)]
    if others:
        found.append((OTHER_PRODUCT, "Security product(s): " + ", ".join(others) + "."))
    if not vba_log_path(workbook, local_app_data, temp).is_file():
        found.append((NO_VBA_LOG, "No VBA log: no macro has run here, or it was blocked before it started."))
    if not cause:
        found.append((NOTHING_IN_OFFICE, "Nothing in Office or in the file explains a block: the endpoint protection is the likely cause."))
    return found


def _findings_section(found: list[tuple[str, str]]) -> list[str]:
    codes = ", ".join(dict.fromkeys(code for code, _ in found))
    return [f"{code}: {text}" for code, text in found] + [f"Codes to quote to the support: {codes}"]


def _file_section(workbook: Path | None) -> list[str]:
    if workbook is None or not workbook.is_file():
        return ["No workbook found. Build one (scripts\\run.ps1 build) or give one with --workbook FILE."]
    found = fingerprint(workbook)
    return [f"Path: {workbook}", name_line(workbook.name), f"Size: {found.size} bytes", f"MD5: {found.md5}", f"SHA-256: {found.sha256}",
            "The MD5 changes with every build: a rule on the file name and folder lasts longer."]


def _origin_section(workbook: Path | None, readings: Readings) -> list[str]:
    if workbook is None or not workbook.is_file():
        return ["Not checked (no workbook)."]
    zone = zone_id(workbook)
    motw = "no" if zone is None else f"yes (ZoneId={zone}, {zone_name(zone)})"
    trusted = "yes" if in_trusted_location(workbook, readings.trusted_locations) else "no"
    return [f"Mark of the Web: {motw}",
            f"Trusted location of Excel: {trusted} ({len(readings.trusted_locations)} location(s) read)"]


def _security_section(products: list[str] | None) -> list[str]:
    if products is None:
        return ["Windows Security Center not available on this machine."]
    if not products:
        return ["none registered in Windows Security Center (an endpoint protection may not register itself there)."]
    return [f"- {name}" for name in products]


def _interpreters_section(interpreters: list[Path]) -> list[str]:
    lines = [f"- {path} (MD5 {fingerprint(path).md5})" for path in interpreters if path.is_file()]
    return lines or ["not listed."]


def _log_section(workbook: Path | None, local_app_data: str, temp: str, count: int) -> tuple[str, list[str]]:
    path = vba_log_path(workbook, local_app_data, temp)
    title = f"VBA log: last {count} lines of {path}"
    if not path.is_file():
        return title, ["no log file: no macro has run on this machine, or it is blocked before it starts."]
    return title, tail(path, count) or ["the log is empty."]


def report(workbook: Path | None, readings: Readings, *, home: str, local_app_data: str, temp: str,
           log_lines: int = 10) -> str:
    """The text to paste into the ticket. Only the VBA log is quoted (it holds no token, URL or Jira answer by design)."""
    log_title, log = _log_section(workbook, local_app_data, temp, log_lines)
    office = [f"{name}: {value}" for name, value in {**readings.office, **readings.policies}.items()] or ["not read."]
    found = findings(workbook, readings, local_app_data, temp)
    sections = [("Findings", _findings_section(found)), ("File", _file_section(workbook)), ("Origin", _origin_section(workbook, readings)), ("Office", office),
                ("Security products", _security_section(readings.security_products)),
                ("Interpreters of the fake Jira", _interpreters_section(readings.interpreters)), (log_title, log)]
    text = [f"jira2dashboard diagnostic, {dt.datetime.now():%Y-%m-%d %H:%M}",
            "Nothing of Config, no token, e-mail, URL or Jira answer is read. Check the text before sharing it."]
    for title, lines in sections:
        text += ["", f"[{title}]", *lines]
    return mask_home("\n".join(text) + "\n", home)
