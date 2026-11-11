<#
.SYNOPSIS
Builds the files of a release in out\release\<version>; with -Publish, creates the GitHub release.

.DESCRIPTION
Needs a clean working tree (everything committed). Builds in out\release\<version>:
  jira2dashboard-<version>.zip  the code of HEAD (git archive; .gitattributes leaves out CLAUDE.md,
                                archive\, .claude\ and .venv\)
  jira2dashboard.xlsm           the workbook to deliver (.\scripts\run.ps1 build; needs Excel)
  notes.md                      the release notes, with the SHA256 of both files
Without -Publish it stops there: a dry run, nothing leaves the machine. With -Publish it creates the GitHub release
and its tag on HEAD, which must already be pushed (gh must be logged in).
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
USE CASES
  Check a release before publishing  .\scripts\release.ps1 -Version v1.0.0            (dry run: nothing leaves the machine)
  Publish it                         .\scripts\release.ps1 -Version v1.0.0 -Publish   (commit and push first: HEAD must be origin/main)
  "the working tree has changes"     commit or stash, then run again
  "HEAD is not origin/main"          push first
  The run stopped half way           fix the cause and run it again: out\release\<version> is rebuilt, nothing is published before the end
  Where is the log                   %LOCALAPPDATA%\jira2dashboard\logs\release.log (last run only)


.PARAMETER Version
The version, like v1.0.0. Required.

.PARAMETER Publish
Also create the release and its tag on GitHub. Outward-facing: the two files become public.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
.\scripts\release.ps1 -Version v1.0.0
Dry run: the zip, the workbook and notes.md in out\release\v1.0.0.

.EXAMPLE
.\scripts\release.ps1 -Version v1.0.0 -Publish
The same, then the GitHub release v1.0.0 with both files.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Version = "",
    [switch]$Publish,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "release: J2D-SCR-001 - unknown argument $($Rest -join ' '). See .\scripts\release.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Log of this run (last run only): %LOCALAPPDATA%\jira2dashboard\logs\release.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $env:LOCALAPPDATA "jira2dashboard\logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "release.log") -Force | Out-Null
        $J2D_LogOwner = "release"
    } catch { }
}

function Fail([string]$code, [string]$message) {
    Write-Host "release: $code - $message"
    exit 1
}

# 1. The release is a commit: a clean tree, and with -Publish that commit is on GitHub
if ($Version -notmatch '^v\d+\.\d+\.\d+$') { Fail "J2D-SCR-010" "the version must look like v1.0.0 (got '$Version')." }
if (git status --porcelain) { Fail "J2D-SCR-011" "the working tree has changes: commit or stash them first." }
$head = (git rev-parse HEAD).Trim()
if ($Publish) {
    gh auth status          # no redirection: in PowerShell 5.1 a redirected stderr becomes an error under "Stop"
    if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-012" "gh is not logged in: run gh auth login." }
    git fetch -q origin main
    if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-013" "git fetch failed." }
    if ($head -ne (git rev-parse origin/main).Trim()) { Fail "J2D-SCR-014" "HEAD is not origin/main: push first." }
    if (git ls-remote --tags origin "refs/tags/$Version") { Fail "J2D-SCR-015" "the tag $Version already exists on GitHub." }
}

# 2. The workbook, built from the same commit. The build empties out\inter first; out\release and the packages of
#    older releases (out\release\<version>) stay.
& (Join-Path $PSScriptRoot "run.ps1") build
if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-016" "the build of the workbook failed (see above)." }
$name = "jira2dashboard-$Version"
$dist = Join-Path $root "out\release\$Version"
New-Item -ItemType Directory -Force $dist | Out-Null
$xlsm = Join-Path $dist "jira2dashboard.xlsm"
Copy-Item (Join-Path $root "out\release\jira2dashboard.xlsm") $xlsm
Write-Host "release: $xlsm"

# 3. The code: git archive of HEAD, then check what it holds
$zip = Join-Path $dist "$name.zip"
git archive --format=zip "--prefix=$name/" -o $zip HEAD
if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-017" "git archive failed." }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($zip)
try { $entries = @($archive.Entries | ForEach-Object { $_.FullName }) } finally { $archive.Dispose() }
$leftOut = @($entries | Where-Object { $_ -match '^[^/]+/(CLAUDE\.md|archive/|\.claude/|\.venv/)' })
if ($leftOut.Count -gt 0) { Fail "J2D-SCR-018" ("the zip holds files it must leave out: " + ($leftOut -join ", ")) }
foreach ($must in "README.md", "scripts/run.ps1", "scripts/setup.ps1", "src/jira2dashboard/build_workbook.py") {
    if ($entries -notcontains "$name/$must") { Fail "J2D-SCR-019" "the zip misses $must." }
}
Write-Host ("release: {0} ({1} entries)" -f $zip, $entries.Count)

# 4. The release notes, with the SHA256 and the MD5 of both files (an allow-list of an endpoint protection often asks for the MD5)
$zipHash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
$xlsmHash = (Get-FileHash $xlsm -Algorithm SHA256).Hash.ToLower()
$zipMd5 = (Get-FileHash $zip -Algorithm MD5).Hash.ToLower()
$xlsmMd5 = (Get-FileHash $xlsm -Algorithm MD5).Hash.ToLower()
$xlsmSize = (Get-Item $xlsm).Length
$notes = @"
jira2dashboard $Version (commit $head)

Files:
- jira2dashboard.xlsm: the workbook, ready to use on Windows with Excel (no Python needed). After the download:
  right-click > Properties > Unblock (or Unblock-File .\jira2dashboard.xlsm), then open it and enable the content.
  Please keep the file name jira2dashboard.xlsm and put it in the same folder on every machine (the one agreed with the
  security team): a new version replaces the file in place, so an allow-list of the endpoint protection on that path or name
  keeps working. A rule on the hash does not (it changes with every build).
- $name.zip: the code, to build the workbook yourself (Windows, Excel, Python 3.14):
  scripts\setup.ps1, then scripts\run.ps1 build.

Setup and use: README.md, section "Livraison".

SHA256:
- jira2dashboard.xlsm: $xlsmHash
- $name.zip: $zipHash

MD5 (for allow-lists that ask for it; it changes with every build, a rule on the file name and folder lasts longer):
- jira2dashboard.xlsm: $xlsmMd5 ($xlsmSize bytes)
- $name.zip: $zipMd5
"@
$notesFile = Join-Path $dist "notes.md"
[System.IO.File]::WriteAllText($notesFile, $notes, (New-Object System.Text.UTF8Encoding $false))

if (-not $Publish) {
    Write-Host "release: dry run done, files in $dist. Run again with -Publish to create the release on GitHub."
    exit 0
}

# 5. Publish: the tag is created on GitHub, on HEAD
gh release create $Version $xlsm $zip --target $head --title "jira2dashboard $Version" --notes-file $notesFile
if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-020" "gh release create failed." }
Write-Host "release: published $Version"
