"""Diagnose a macro block: a report to paste into the ticket for the team of the endpoint protection.

Describes the workbook (size, MD5 and SHA-256: an allow-list often asks for the MD5), where it comes from (Mark of the Web,
Trusted Location of Excel), Office and its macro policies, the security products registered in Windows Security Center,
the interpreters of the fake Jira and the last lines of the VBA log. It never opens the workbook (no macro, no Config sheet),
never reads python.log nor the Windows event log, and masks the home folder of the user. Check the text before sharing it.
Default workbook: out\\release\\jira2dashboard.xlsm, then the manual-test copy.
"""
import os
import subprocess
import sys
import winreg
from pathlib import Path

from jira2dashboard import diagnose as D
from jira2dashboard.cli.args import parser
from jira2dashboard.paths import DELIVERED_NAME, MANUAL_TEST_DIR, RELEASE_DIR

OFFICE_KEY = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
EXCEL_SECURITY = r"Software\Microsoft\Office\16.0\Excel\Security"
EXCEL_POLICY = r"Software\Policies\Microsoft\Office\16.0\excel\security"
SECURITY_CENTER = ("Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct "
                   "| ForEach-Object { $_.displayName }")


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 diagnose", __doc__)
    p.add_argument("--workbook", type=Path, help="the workbook to describe (default: the delivered workbook)")
    return p.parse_args(argv)


def default_workbook() -> Path | None:
    for path in (RELEASE_DIR / DELIVERED_NAME, MANUAL_TEST_DIR / "jira2dashboard4manualtest.xlsm"):
        if path.is_file():
            return path
    return None


def registry_value(hive, key: str, name: str):
    try:
        with winreg.OpenKey(hive, key) as opened:
            return winreg.QueryValueEx(opened, name)[0]
    except OSError:
        return None


def office() -> dict[str, str]:
    found = {"Version": registry_value(winreg.HKEY_LOCAL_MACHINE, OFFICE_KEY, "VersionToReport"),
             "Platform": registry_value(winreg.HKEY_LOCAL_MACHINE, OFFICE_KEY, "Platform")}
    found = {name: str(value) for name, value in found.items() if value is not None}
    return found or {"Version": "not found (no Click-to-Run configuration)"}


def policies() -> dict[str, str]:
    where = {D.VBA_USER: (EXCEL_SECURITY, "VBAWarnings"),
             D.VBA_POLICY: (EXCEL_POLICY, "VBAWarnings"),
             D.BLOCK_INTERNET: (EXCEL_POLICY, "BlockContentExecutionFromInternet")}
    values = {label: registry_value(winreg.HKEY_CURRENT_USER, key, name) for label, (key, name) in where.items()}
    return {label: "not set" if value is None else str(value) for label, value in values.items()}


def trusted_locations() -> list[tuple[str, bool]]:
    found = []
    for root in (EXCEL_SECURITY + r"\Trusted Locations", EXCEL_POLICY + r"\trustedlocations"):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, root) as parent:
                names = [winreg.EnumKey(parent, n) for n in range(winreg.QueryInfoKey(parent)[0])]
        except OSError:
            continue
        for name in names:
            path = registry_value(winreg.HKEY_CURRENT_USER, f"{root}\\{name}", "Path")
            if path:
                allow = registry_value(winreg.HKEY_CURRENT_USER, f"{root}\\{name}", "AllowSubfolders")
                found.append((os.path.expandvars(str(path)), bool(allow)))
    return found


def security_products() -> list[str] | None:
    """The products registered in Windows Security Center; None when it cannot be asked (a server, a timeout)."""
    try:
        done = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", SECURITY_CENTER],
                              capture_output=True, text=True, encoding="oem", errors="replace", timeout=20)
    except (OSError, subprocess.TimeoutExpired):
        return None
    if done.returncode != 0:
        return None
    return sorted({line.strip() for line in done.stdout.splitlines() if line.strip()})


def interpreters() -> list[Path]:
    """The Python that runs this command and the interpreter behind it (a venv python.exe is only a launcher)."""
    found = []
    for path in (Path(sys.executable), Path(getattr(sys, "_base_executable", sys.executable))):
        if path not in found:
            found.append(path)
    return found


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    readings = D.Readings(office=office(), policies=policies(), trusted_locations=trusted_locations(),
                          security_products=security_products(), interpreters=interpreters())
    sys.stdout.reconfigure(errors="replace")
    print(D.report(args.workbook or default_workbook(), readings, home=str(Path.home()),
                   local_app_data=os.environ.get("LOCALAPPDATA", ""), temp=os.environ.get("TEMP", "")), end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())
