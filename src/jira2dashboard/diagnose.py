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
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return []
    return lines[-count:] if count > 0 else []


def vba_log_path(local_app_data: str, temp: str) -> Path:
    """Where modLog writes: %LOCALAPPDATA%\\jira2dashboard\\logs, or the TEMP folder when LOCALAPPDATA is empty."""
    return Path(local_app_data or temp) / "jira2dashboard" / "logs" / "jira2dashboard.log"


def _file_section(workbook: Path | None) -> list[str]:
    if workbook is None or not workbook.is_file():
        return ["No workbook found. Build one (scripts\\run.ps1 build) or give one with --workbook FILE."]
    found = fingerprint(workbook)
    return [f"Path: {workbook}", f"Size: {found.size} bytes", f"MD5: {found.md5}", f"SHA-256: {found.sha256}",
            "The MD5 changes with every build: a rule on the path or on a signature lasts longer."]


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


def _log_section(local_app_data: str, temp: str, count: int) -> tuple[str, list[str]]:
    path = vba_log_path(local_app_data, temp)
    title = f"VBA log: last {count} lines of {path}"
    if not path.is_file():
        return title, ["no log file: no macro has run on this machine, or it is blocked before it starts."]
    return title, tail(path, count) or ["the log is empty."]


def report(workbook: Path | None, readings: Readings, *, home: str, local_app_data: str, temp: str,
           log_lines: int = 10) -> str:
    """The text to paste into the ticket. Only the VBA log is quoted (it holds no token, URL or Jira answer by design)."""
    log_title, log = _log_section(local_app_data, temp, log_lines)
    office = [f"{name}: {value}" for name, value in {**readings.office, **readings.policies}.items()] or ["not read."]
    sections = [("File", _file_section(workbook)), ("Origin", _origin_section(workbook, readings)), ("Office", office),
                ("Security products", _security_section(readings.security_products)),
                ("Interpreters of the fake Jira", _interpreters_section(readings.interpreters)), (log_title, log)]
    text = [f"jira2dashboard diagnostic, {dt.datetime.now():%Y-%m-%d %H:%M}",
            "Nothing of Config, no token, e-mail, URL or Jira answer is read. Check the text before sharing it."]
    for title, lines in sections:
        text += ["", f"[{title}]", *lines]
    return mask_home("\n".join(text) + "\n", home)
