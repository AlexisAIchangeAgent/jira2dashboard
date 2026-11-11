# Pester 3.4 tests of scripts\*.ps1, run IN PROCESS so that Invoke-Pester -CodeCoverage can measure them
# (tests\ps\Run-ScriptTests.ps1 does it: .\scripts\run.ps1 test-scripts).
#
# Each script runs from a sandbox copy (the script finds its root from its own folder): no .venv, no git, no gh, no Excel, no
# network and nothing of the real repository is touched. Native programs are replaced by functions of the same name
# (python.exe by the function named like its full path): they note their call in $global:FakeCalls and answer what the test
# scripted. Start-Process, Start-Sleep and Get-Date are mocked; the TCP port checks of start-mock.ps1 are real (a listener of the test).
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = (Resolve-Path (Join-Path $here "..\..")).Path
$global:Saved = @{}
foreach ($name in "LOCALAPPDATA", "TEMP", "PYTHONPATH", "PYTHONIOENCODING", "J2D_ACTIVATED") {
    $global:Saved[$name] = [Environment]::GetEnvironmentVariable($name, "Process")
}
$global:StartDir = (Get-Location).Path

if (-not $env:J2D_PS_BASE) {                          # run by hand, without the runner: the sandbox copy is made here
    $env:J2D_PS_BASE = Join-Path $repo ("out\inter\ps_" + [guid]::NewGuid().ToString("N").Substring(0, 8))   # long path: see Run-ScriptTests.ps1
    New-Item -ItemType Directory -Force (Join-Path $env:J2D_PS_BASE "scripts") | Out-Null
    Copy-Item (Join-Path $repo "scripts\*.ps1") (Join-Path $env:J2D_PS_BASE "scripts")
}
$global:Base = $env:J2D_PS_BASE
$global:Scripts = Join-Path $global:Base "scripts"
$global:Python = Join-Path $global:Base ".venv\Scripts\python.exe"
$global:FakeCalls = New-Object System.Collections.ArrayList
$global:FakeBehavior = @{}
$global:FakeNames = New-Object System.Collections.ArrayList
$global:Listener = $null

function global:Invoke-Fake([string]$key, $argv) {
    [void]$global:FakeCalls.Add((("{0} {1}" -f $key, ($argv -join " ")).Trim()))
    $global:LASTEXITCODE = 0
    $behavior = $global:FakeBehavior[$key]
    $argList = @($argv)
    if ($behavior) { & $behavior @argList }
}

function Set-Fake([string]$key, [scriptblock]$behavior = $null, [string]$name = "") {
    # A function named $name (a command name, or a full path) that stands for the program; its behavior gets the arguments.
    if (-not $name) { $name = $key }
    $global:FakeBehavior[$key] = $behavior
    New-Item -Path "Function:\global:$name" -Value ([scriptblock]::Create("Invoke-Fake '$key' `$args")) -Force | Out-Null
    [void]$global:FakeNames.Add($name)
}

function Calls { @($global:FakeCalls) }

function Reset-Sandbox {
    Remove-Item (Join-Path $global:Scripts "logs") -Recurse -Force -ErrorAction SilentlyContinue      # the logs of the scripts of the sandbox
    Get-ChildItem Function: | Where-Object { $_.Name -like ($global:Scripts + "*") } | ForEach-Object { Remove-Item "Function:\$($_.Name)" -ErrorAction SilentlyContinue }
    Get-ChildItem $global:Base -Force | Where-Object { $_.Name -ne "scripts" } | Remove-Item -Recurse -Force
    New-Item -ItemType Directory -Force (Join-Path $global:Base "local"), (Join-Path $global:Base "temp") | Out-Null
    $env:LOCALAPPDATA = Join-Path $global:Base "local"
    $env:TEMP = Join-Path $global:Base "temp"
    $global:FakeCalls.Clear()
    $global:FakeBehavior.Clear()
    $global:LASTEXITCODE = 0
    Set-Location $global:Base
}

function Restore-Sandbox {
    for ($i = 0; $i -lt 5; $i++) { try { Stop-Transcript | Out-Null } catch { break } }
    foreach ($name in $global:FakeNames) { Remove-Item "Function:\global:$name" -ErrorAction SilentlyContinue }
    $global:FakeNames.Clear()
    if ($global:Listener) { $global:Listener.Stop(); $global:Listener = $null }
    foreach ($name in $global:Saved.Keys) { [Environment]::SetEnvironmentVariable($name, $global:Saved[$name], "Process") }
    Set-Location $global:StartDir
}

function New-Venv {
    New-Item -ItemType Directory -Force (Join-Path $global:Base ".venv\Scripts") | Out-Null
    Set-Content (Join-Path $global:Base ".venv\Scripts\python.exe") "placeholder: the call goes to the fake function"
    Set-Fake "python" -name $global:Python
}

function Invoke-Script([string]$name, [string[]]$Positional = @(), [hashtable]$Named = @{}) {
    $global:LASTEXITCODE = 0
    $error = ""
    $out = ""
    try { $out = (& (Join-Path $global:Scripts $name) @Positional @Named *>&1 | Out-String) } catch { $error = $_.Exception.Message }
    if ($error) { $out += "`nERROR: $error" }
    [pscustomobject]@{ Out = $out; Code = $global:LASTEXITCODE; Error = $error }
}

function Free-Port {
    $l = New-Object System.Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, 0)
    $l.Start(); $port = $l.LocalEndpoint.Port; $l.Stop()
    $port
}

function Open-Port([int]$port) {
    $global:Listener = New-Object System.Net.Sockets.TcpListener ([Net.IPAddress]::Loopback, $port)
    $global:Listener.Start()
}

function Write-PidFile([int]$port, [long]$start) {
    $folder = Join-Path $env:TEMP "jira2dashboard"
    New-Item -ItemType Directory -Force $folder | Out-Null
    Set-Content (Join-Path $folder "mock-jira-$port.pid") ("{0} {1}" -f $PID, $start) -Encoding Ascii
}

function Own-Start { (Get-Process -Id $PID).StartTime.ToFileTimeUtc() }
function Pid-File([int]$port) { Join-Path (Join-Path $env:TEMP "jira2dashboard") "mock-jira-$port.pid" }
function Log-File([string]$name) { Join-Path $global:Scripts "logs\$name.log" }
function Block-Logs {
    # a file where the folder scripts\logs should be: the script must go on without a log
    Remove-Item (Join-Path $global:Scripts "logs") -Recurse -Force -ErrorAction SilentlyContinue
    Set-Content (Join-Path $global:Scripts "logs") "a file where the log folder should be"
}

# --- the scripts ask for help the same way ---------------------------------------------------------------------------------

Describe "help of every script" {
    BeforeEach { Reset-Sandbox }
    AfterEach { Restore-Sandbox }

    foreach ($script in "run", "setup", "release", "start-mock", "stop-mock", "manual-test", "scenario") {
        foreach ($flag in "-Help", "--help") {
            It "$script.ps1 $flag shows its help and runs nothing" {
                $r = if ($flag -eq "-Help" -and $script -ne "run") { Invoke-Script "$script.ps1" -Named @{ Help = $true } } else { Invoke-Script "$script.ps1" -Positional @($flag) }
                $r.Code | Should Be 0
                $r.Out | Should Match "(?i)(SYNOPSIS|R.SUM.|Usage:)"
                (Calls).Count | Should Be 0
            }
        }
        It "$script.ps1 has a use cases section in its help" {
            $r = if ($script -eq "run") { Invoke-Script "run.ps1" -Positional @("--help") } else { Invoke-Script "$script.ps1" -Named @{ Help = $true } }
            $r.Out | Should Match "(?i)USE CASES"
        }
    }
    foreach ($script in "setup", "release", "start-mock", "stop-mock", "manual-test", "scenario") {
        It "$script.ps1 refuses an unknown argument with exit code 2" {
            $r = Invoke-Script "$script.ps1" -Positional @("--nonsense")
            $r.Code | Should Be 2
            $r.Out | Should Match "J2D-SCR-001 - unknown argument --nonsense"
        }
    }
    It "activate.ps1 shows its help when dot-sourced with --help" {
        $text = (. (Join-Path $global:Scripts "activate.ps1") --help *>&1 | Out-String)
        $text | Should Match "(?i)(SYNOPSIS|R.SUM.)"
        $text | Should Match "USE CASES"
    }
}

# --- run.ps1 ------------------------------------------------------------------------------------------------------------

Describe "run.ps1" {
    BeforeEach { Reset-Sandbox }
    AfterEach { Restore-Sandbox }

    It "without a command lists the commands and the use cases and exits with 2" {
        $r = Invoke-Script "run.ps1"
        $r.Code | Should Be 2
        $r.Out | Should Match "Usage: .\\scripts\\run.ps1 <command>"
        $r.Out | Should Match "test-scripts"
        $r.Out | Should Match "Use cases:"
    }
    It "test-scripts runs the Pester runner of tests\\ps with the arguments untouched" {
        New-Venv
        Set-Fake "powershell"
        $r = Invoke-Script "run.ps1" -Positional @("test-scripts", "-Minimum", "97")
        $r.Code | Should Be 0
        (Calls) | Should Be @("powershell -NoProfile -ExecutionPolicy Bypass -File " + (Join-Path $global:Base "tests\ps\Run-ScriptTests.ps1") + " -Minimum 97")
    }
    foreach ($flag in "--help", "-h", "-help", "help", "/?") {
        It "$flag lists the commands and exits with 0" {
            $r = Invoke-Script "run.ps1" -Positional @($flag)
            $r.Code | Should Be 0
            $r.Out | Should Match "vba-tests"
        }
    }
    It "refuses an unknown command with 2" {
        $r = Invoke-Script "run.ps1" -Positional @("nonsense")
        $r.Code | Should Be 2
        $r.Out | Should Match "J2D-SCR-002 - unknown command 'nonsense'"
    }
    It "asks for setup.ps1 when there is no .venv" {
        $r = Invoke-Script "run.ps1" -Positional @("test")
        $r.Error | Should Match "J2D-SCR-003 - No .venv: run scripts.setup.ps1 -Dev first"
    }

    $commands = [ordered]@{
        "test" = "-m pytest"; "coverage" = "-m pytest --cov --cov-report=term-missing"
        "test-excel" = "-m pytest -m excel and not network"; "test-network" = "-m pytest -m excel and network"
        "vba-tests" = "-m jira2dashboard.testbench.run_vba_tests"; "check-machine" = "-m jira2dashboard.cli.check_machine"
        "diagnose" = "-m jira2dashboard.cli.diagnose"; "mock-jira" = "-m jira2dashboard.mockjira"
        "mock-workbook" = "-m jira2dashboard.mock_workbook"; "build" = "-m jira2dashboard.build_workbook"
        "clean" = "-m jira2dashboard.clean"; "extract" = "-m jira2dashboard.standalone"
    }
    foreach ($command in $commands.Keys) {
        It "$command runs python $($commands[$command]) with the arguments untouched" {
            New-Venv
            $r = Invoke-Script "run.ps1" -Positional @($command, "--out", "x.xlsm", "-k", "a b")
            $r.Code | Should Be 0
            (Calls) | Should Be @("python $($commands[$command]) --out x.xlsm -k a b")
        }
    }
    It "passes the exit code of the command back" {
        New-Venv
        Set-Fake "python" { $global:LASTEXITCODE = 3 } -name $global:Python
        (Invoke-Script "run.ps1" -Positional @("test")).Code | Should Be 3
    }
    It "sets PYTHONPATH to src and the UTF-8 output for the command" {
        New-Venv
        $global:Seen = ""
        Set-Fake "python" { $global:Seen = "$env:PYTHONPATH|$env:PYTHONIOENCODING" } -name $global:Python
        Invoke-Script "run.ps1" -Positional @("test") | Out-Null
        $global:Seen | Should Be ((Join-Path $global:Base "src") + "|utf-8")
    }
    It "writes the log of the run, which is its transcript, in a file of its command" {
        New-Venv
        Invoke-Script "run.ps1" -Positional @("test") | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "run_test") | Should Be $true
        Test-Path (Log-File "run") | Should Be $false
    }
    It "keeps the output of python in the log, its error output too, and its exit code" {
        New-Venv
        Set-Fake "python" {                                     # a real native program writes the error line, as python.exe does
            "a line of python"; & cmd.exe /c "echo a line on stderr 1>&2"; $global:LASTEXITCODE = 4
        } -name $global:Python
        $r = Invoke-Script "run.ps1" -Positional @("build")
        Stop-Transcript | Out-Null
        $r.Code | Should Be 4
        $r.Out | Should Match "a line of python"
        $log = Get-Content (Log-File "run_build") -Raw
        $log | Should Match "a line of python"
        $log | Should Match "a line on stderr"           # (here the fake is a function: PowerShell also logs it as an error record;
                                                          # python.exe called with 2>&1 gives the plain line only)
    }
    It "sets the output of python unbuffered, and gives the console its encoding back" {
        New-Venv
        $before = [Console]::OutputEncoding.WebName
        $global:Seen = ""
        Set-Fake "python" { $global:Seen = "$env:PYTHONUNBUFFERED|$([Console]::OutputEncoding.WebName)" } -name $global:Python
        Invoke-Script "run.ps1" -Positional @("test") | Out-Null
        $global:Seen | Should Be "1|utf-8"
        [Console]::OutputEncoding.WebName | Should Be $before
    }
    It "goes on when the log cannot be written" {
        New-Venv
        Block-Logs
        $r = Invoke-Script "run.ps1" -Positional @("test")
        $r.Code | Should Be 0
        (Calls).Count | Should Be 1
    }
}

# --- setup.ps1 ----------------------------------------------------------------------------------------------------------

Describe "setup.ps1" {
    BeforeEach { Reset-Sandbox }
    AfterEach { Restore-Sandbox }

    function Fake-Py {
        Set-Fake "py" {
            New-Item -ItemType Directory -Force (Join-Path $global:Base ".venv\Scripts") | Out-Null
            Set-Content (Join-Path $global:Base ".venv\Scripts\python.exe") "placeholder"
            Set-Fake "python" -name $global:Python           # again: the file exists now, and the call must still reach the function
        }
    }

    It "creates .venv with Python 3.14, upgrades pip and installs requirements.txt" {
        Fake-Py
        $r = Invoke-Script "setup.ps1"
        $r.Code | Should Be 0
        $r.Out | Should Match "Setup OK \(requirements.txt\)"
        (Calls) | Should Be @("py -3.14 -m venv .venv", "python -m pip install --upgrade pip",
                              "python -m pip install -r requirements.txt", "python -c import win32com.client, pythoncom; print('pywin32 import OK')")
    }
    It "-Dev installs requirements-dev.txt" {
        Fake-Py
        (Invoke-Script "setup.ps1" -Named @{ Dev = $true }).Out | Should Match "Setup OK \(requirements-dev.txt\)"
        (Calls) -contains "python -m pip install -r requirements-dev.txt" | Should Be $true
    }
    It "-Proxy goes to both pip calls" {
        Fake-Py
        Invoke-Script "setup.ps1" -Named @{ Proxy = "http://proxy.example:8080" } | Out-Null
        @(Calls | Where-Object { $_ -match "--proxy http://proxy.example:8080" }).Count | Should Be 2
    }
    It "keeps an existing .venv: py is not called" {
        New-Venv
        (Invoke-Script "setup.ps1").Out | Should Match "Setup OK"
        @(Calls | Where-Object { $_ -like "py *" }).Count | Should Be 0
    }
    It "stops when the venv cannot be created" {
        Set-Fake "py" { $global:LASTEXITCODE = 1 }
        (Invoke-Script "setup.ps1").Error | Should Match "py -3.14 -m venv failed"
    }
    It "stops when pip cannot be upgraded" {
        New-Venv
        Set-Fake "python" { if ($args -contains "--upgrade") { $global:LASTEXITCODE = 1 } } -name $global:Python
        (Invoke-Script "setup.ps1").Error | Should Match "pip upgrade failed"
    }
    It "stops when the requirements cannot be installed" {
        New-Venv
        Set-Fake "python" { if ($args -contains "-r") { $global:LASTEXITCODE = 1 } } -name $global:Python
        (Invoke-Script "setup.ps1").Error | Should Match "pip install -r requirements.txt failed"
    }
    It "stops when pywin32 cannot be imported" {
        New-Venv
        Set-Fake "python" { if ($args -contains "-c") { $global:LASTEXITCODE = 1 } } -name $global:Python
        (Invoke-Script "setup.ps1").Error | Should Match "pywin32 cannot be imported"
    }
    It "writes its log, and goes on when the log cannot be written" {
        New-Venv
        Invoke-Script "setup.ps1" | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "setup") | Should Be $true
        Block-Logs
        (Invoke-Script "setup.ps1").Out | Should Match "Setup OK"
    }
}

# --- stop-mock.ps1 ------------------------------------------------------------------------------------------------------

Describe "stop-mock.ps1" {
    BeforeEach { Reset-Sandbox; Set-Fake "taskkill" }
    AfterEach { Restore-Sandbox }

    It "says so when no fake Jira was started by the scripts" {
        $r = Invoke-Script "stop-mock.ps1"
        $r.Code | Should Be 0
        $r.Out | Should Match "no fake Jira started by these scripts on port 8099"
        (Calls).Count | Should Be 0
    }
    It "kills the process tree of the window it started (same pid and start time) and forgets the pid file" {
        Write-PidFile 8099 (Own-Start)
        $r = Invoke-Script "stop-mock.ps1"
        $r.Out | Should Match "fake Jira of port 8099 stopped"
        (Calls) | Should Be @("taskkill /PID $PID /T /F")
        Test-Path (Pid-File 8099) | Should Be $false
    }
    It "leaves alone a program that reused the pid" {
        Write-PidFile 8099 ((Own-Start) + 1)
        (Invoke-Script "stop-mock.ps1").Out | Should Match "had already stopped"
        (Calls).Count | Should Be 0
        Test-Path (Pid-File 8099) | Should Be $false
    }
    It "tells when the process is gone" {
        $folder = Join-Path $env:TEMP "jira2dashboard"
        New-Item -ItemType Directory -Force $folder | Out-Null
        Set-Content (Join-Path $folder "mock-jira-8100.pid") "999999 1" -Encoding Ascii
        (Invoke-Script "stop-mock.ps1" -Named @{ Port = 8100 }).Out | Should Match "port 8100 had already stopped"
    }
    It "writes its log, and goes on when the log cannot be written" {
        Invoke-Script "stop-mock.ps1" | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "stop-mock") | Should Be $true
        Block-Logs
        (Invoke-Script "stop-mock.ps1").Code | Should Be 0
    }
}

# --- start-mock.ps1 -----------------------------------------------------------------------------------------------------

Describe "start-mock.ps1" {
    BeforeEach {
        Reset-Sandbox
        Set-Fake "taskkill"
        Mock Start-Sleep { }
        $global:Port = Free-Port
        $global:Window = [pscustomobject]@{ Id = $PID; StartTime = (Get-Process -Id $PID).StartTime; HasExited = $false }
    }
    AfterEach { Restore-Sandbox }

    function Fake-Server([bool]$listen) {
        # Start-Process of the script: the "server" starts listening (or not).
        Mock Start-Process { if ($global:Listen) { Open-Port $global:Port }; $global:Window }
        $global:Listen = $listen
    }

    It "starts the fake Jira, waits until it answers, and names the interpreter that serves it" {
        Set-Fake "netstat" { "  TCP    127.0.0.1:$($global:Port)    0.0.0.0:0    LISTENING    $PID" }
        Fake-Server $true
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port; Issues = 150 }
        $r.Code | Should Be 0
        $r.Out | Should Match "fake Jira on http://127.0.0.1:$($global:Port) \(150 issues and the known anomalies\)"
        $r.Out | Should Match "(?s)start-mock: python: .*powershell.*MD5 [0-9a-f]{32},\s+SHA256 [0-9a-f]{64}"
        $r.Out | Should Match "stop it with .\\scripts\\stop-mock.ps1 -Port $($global:Port)"
        Test-Path (Pid-File $global:Port) | Should Be $true
        Assert-MockCalled Start-Process -Scope It -Times 1 -ParameterFilter { $WindowStyle -eq "Minimized" -and ($ArgumentList -contains "--anomalies") -and ($ArgumentList -contains "mock-jira") }
    }
    It "-Hidden starts it with no window" {
        Set-Fake "netstat"
        Fake-Server $true
        Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port; Hidden = $true } | Out-Null
        Assert-MockCalled Start-Process -Scope It -Times 1 -ParameterFilter { $WindowStyle -eq "Hidden" }
    }
    It "says the interpreter was not found, without failing, when no listener has a path" {
        Set-Fake "netstat" { "nothing listens here" }
        Fake-Server $true
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }
        $r.Code | Should Be 0
        $r.Out | Should Match "python: not found \(not an error\)"
    }
    It "does not start a second one: the window it started still runs" {
        Write-PidFile $global:Port (Own-Start)
        Fake-Server $false
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }
        $r.Code | Should Be 0
        $r.Out | Should Match "a fake Jira started by these scripts already runs"
        Assert-MockCalled Start-Process -Scope It -Times 0
    }
    It "refuses a port held by another program and leaves it alone" {
        Open-Port $global:Port
        Fake-Server $false
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }
        $r.Code | Should Be 1
        $r.Out | Should Match "J2D-SCR-030 - port \d+ is held by another program"
        Assert-MockCalled Start-Process -Scope It -Times 0
    }
    It "fails at once when the fake Jira stops at once" {
        $global:Window = [pscustomobject]@{ Id = $PID; StartTime = (Get-Process -Id $PID).StartTime; HasExited = $true }
        Fake-Server $false
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }
        $r.Code | Should Be 1
        $r.Out | Should Match "J2D-SCR-031 - the fake Jira stopped at once"
        Test-Path (Pid-File $global:Port) | Should Be $false
    }
    It "stops what it started and fails when nothing answers after 60 seconds" {
        $global:Dates = 0
        Mock Get-Date { $global:Dates++; if ($global:Dates -le 1) { [datetime]::Now } else { [datetime]::Now.AddMinutes(5) } }
        Fake-Server $false
        $r = Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }
        $r.Code | Should Be 1
        $r.Out | Should Match "J2D-SCR-032 - no answer on port $($global:Port) after 60 s"
        (Calls) -contains "taskkill /PID $PID /T /F" | Should Be $true
    }
    It "writes its log, and goes on when the log cannot be written" {
        Write-PidFile $global:Port (Own-Start)
        Fake-Server $false
        Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port } | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "start-mock") | Should Be $true
        Block-Logs
        (Invoke-Script "start-mock.ps1" -Named @{ Port = $global:Port }).Code | Should Be 0
    }
}

# --- release.ps1 --------------------------------------------------------------------------------------------------------

Describe "release.ps1" {
    BeforeEach {
        Reset-Sandbox
        New-Venv
        $global:Dirty = ""; $global:Origin = "abc123"; $global:Tag = ""; $global:FetchExit = 0
        $global:GhAuthExit = 0; $global:GhCreateExit = 0; $global:BuildExit = 0; $global:ArchiveExit = 0
        $global:ZipNames = "README.md", "scripts/run.ps1", "scripts/setup.ps1", "src/jira2dashboard/build_workbook.py", "scripts/release.ps1"
        Set-Fake "python" {
            if ($args -contains "jira2dashboard.build_workbook") {
                New-Item -ItemType Directory -Force "out\release" | Out-Null
                Set-Content "out\release\jira2dashboard.xlsm" "workbook"
                $global:LASTEXITCODE = $global:BuildExit
            }
        } -name $global:Python
        Set-Fake "git" {
            $line = $args -join " "
            if ($line -like "status --porcelain*") { $global:Dirty }
            elseif ($line -eq "rev-parse HEAD") { "abc123" }
            elseif ($line -eq "rev-parse origin/main") { $global:Origin }
            elseif ($line -like "fetch*") { $global:LASTEXITCODE = $global:FetchExit }
            elseif ($line -like "ls-remote*") { $global:Tag }
            elseif ($line -like "archive*") {
                $global:LASTEXITCODE = $global:ArchiveExit
                if ($global:ArchiveExit -eq 0) {
                    $zip = $args[[array]::IndexOf($args, "-o") + 1]
                    $prefix = ($args | Where-Object { $_ -like "--prefix=*" }) -replace "^--prefix=", ""
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    $archive = [IO.Compression.ZipFile]::Open($zip, "Create")
                    foreach ($entry in $global:ZipNames) {
                        $writer = New-Object IO.StreamWriter ($archive.CreateEntry($prefix + $entry).Open()); $writer.Write("x"); $writer.Dispose()
                    }
                    $archive.Dispose()
                }
            }
        }
        Set-Fake "gh" {
            if ($args[0] -eq "auth") { "Logged in"; $global:LASTEXITCODE = $global:GhAuthExit }
            elseif ($args[0] -eq "release") { $global:LASTEXITCODE = $global:GhCreateExit }
        }
        Copy-Item (Join-Path $global:Scripts "run.ps1") $global:Scripts -ErrorAction SilentlyContinue
    }
    AfterEach { Restore-Sandbox }

    It "refuses a version that does not look like v1.0.0" {
        foreach ($bad in "", "1.0.0", "v1.0", "v0.0.0-dryrun") {
            $r = Invoke-Script "release.ps1" -Named @{ Version = $bad }
            $r.Code | Should Be 1
            $r.Out | Should Match "J2D-SCR-010 - the version must look like v1.0.0"
        }
    }
    It "refuses a working tree with changes" {
        $global:Dirty = " M README.md"
        $r = Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }
        $r.Code | Should Be 1
        $r.Out | Should Match "J2D-SCR-011 - the working tree has changes"
    }
    It "dry run: builds the workbook and the zip, writes the notes with the hashes, publishes nothing" {
        $r = Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }
        $r.Code | Should Be 0
        $r.Out | Should Match "dry run done"
        $dist = Join-Path $global:Base "out\release\v1.0.0"
        Test-Path (Join-Path $dist "jira2dashboard.xlsm") | Should Be $true
        Test-Path (Join-Path $dist "jira2dashboard-v1.0.0.zip") | Should Be $true
        $notes = Get-Content (Join-Path $dist "notes.md") -Raw
        $notes | Should Match "jira2dashboard v1.0.0 \(commit abc123\)"
        (Calls) -contains "python -m jira2dashboard.build_workbook --version v1.0.0" | Should Be $true     # the version goes into the workbook
        $notes | Should Match "SHA256:\s+- jira2dashboard.xlsm: [0-9a-f]{64}"
        $notes | Should Match "MD5 .*\s+- jira2dashboard.xlsm: [0-9a-f]{32} \(\d+ bytes\)"
        @(Calls | Where-Object { $_ -like "gh *" }).Count | Should Be 0
    }
    It "stops when the build of the workbook fails" {
        $global:BuildExit = 1
        $r = Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }
        $r.Code | Should Be 1
        $r.Out | Should Match "the build of the workbook failed"
    }
    It "stops when git archive fails" {
        $global:ArchiveExit = 1
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }).Out | Should Match "git archive failed"
    }
    It "refuses a zip that holds a file that must stay out" {
        $global:ZipNames += "CLAUDE.md"
        $r = Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }
        $r.Code | Should Be 1
        $r.Out | Should Match "the zip holds files it must leave out: .*CLAUDE.md"
    }
    It "refuses a zip that misses a file of the code" {
        $global:ZipNames = "README.md", "scripts/run.ps1"
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" }).Out | Should Match "the zip misses scripts/setup.ps1"
    }
    It "-Publish creates the release on HEAD with the workbook, the zip and the notes" {
        $r = Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }
        $r.Code | Should Be 0
        $r.Out | Should Match "published v1.0.0"
        $create = @(Calls | Where-Object { $_ -like "gh release create*" })
        $create.Count | Should Be 1
        $create[0] | Should Match "gh release create v1.0.0 .*jira2dashboard.xlsm .*jira2dashboard-v1.0.0.zip --target abc123 --title jira2dashboard v1.0.0 --notes-file .*notes.md"
    }
    It "-Publish needs gh logged in" {
        $global:GhAuthExit = 1
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }).Out | Should Match "J2D-SCR-012 - gh is not logged in"
    }
    It "-Publish needs git fetch to work" {
        $global:FetchExit = 1
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }).Out | Should Match "git fetch failed"
    }
    It "-Publish needs HEAD to be origin/main" {
        $global:Origin = "other"
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }).Out | Should Match "HEAD is not origin/main: push first"
    }
    It "-Publish refuses a tag that already exists" {
        $global:Tag = "abc123 refs/tags/v1.0.0"
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }).Out | Should Match "the tag v1.0.0 already exists on GitHub"
    }
    It "-Publish reports a failure of gh release create" {
        $global:GhCreateExit = 1
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0"; Publish = $true }).Out | Should Match "gh release create failed"
    }
    It "writes its log, and goes on when the log cannot be written" {
        Invoke-Script "release.ps1" -Named @{ Version = "v1.0.0" } | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "release") | Should Be $true
        Block-Logs
        (Invoke-Script "release.ps1" -Named @{ Version = "v1.0.1" }).Code | Should Be 0
    }
}

# --- manual-test.ps1 ----------------------------------------------------------------------------------------------------

Describe "manual-test.ps1" {
    BeforeEach {
        Reset-Sandbox
        Mock Start-Process { }
        $global:StopExit = 0; $global:BuildExit = 0; $global:MockWbExit = 0; $global:StartExit = 0; $global:ExtractExit = 0
        # the scripts it calls are replaced by functions named like their full path: this test is about the orchestration
        Set-Fake "stop-mock" { $global:LASTEXITCODE = $global:StopExit } -name (Join-Path $global:Scripts "stop-mock.ps1")
        Set-Fake "start-mock" { $global:LASTEXITCODE = $global:StartExit } -name (Join-Path $global:Scripts "start-mock.ps1")
        Set-Fake "run" {
            if ($args[0] -eq "build") { New-Item -ItemType Directory -Force "out\release" | Out-Null; Set-Content "out\release\jira2dashboard.xlsm" "wb"; $global:LASTEXITCODE = $global:BuildExit }
            if ($args[0] -eq "mock-workbook") { $global:LASTEXITCODE = $global:MockWbExit }
            if ($args[0] -eq "extract") { $global:LASTEXITCODE = $global:ExtractExit }
        } -name (Join-Path $global:Scripts "run.ps1")
    }
    AfterEach { Restore-Sandbox }

    It "stops the old fake Jira, builds the workbook when missing, prepares the copy, starts the fake Jira and opens Excel" {
        $r = Invoke-Script "manual-test.ps1"
        $r.Code | Should Be 0
        (Calls | ForEach-Object { ($_ -split " ")[0] }) -join "," | Should Be "stop-mock,run,run,start-mock"
        (Calls)[0] | Should Be "stop-mock -Port 8099"
        (Calls)[1] | Should Be "run build"
        (Calls)[2] | Should Match "^run mock-workbook --workbook .*out.release.jira2dashboard.xlsm --port 8099$"
        $r.Out | Should Match "1. Read the Jira configuration"
        $r.Out | Should Match "2. Extract from Jira and compute"
        $r.Out | Should Match "Excel may crash right after the last search page"
        Assert-MockCalled Start-Process -Scope It -Times 1 -ParameterFilter { $FilePath -eq "excel.exe" -and ($ArgumentList -contains "/x") -and ($ArgumentList -join " ") -match "jira2dashboard4manualtest.xlsm" }
    }
    It "does not build when out\release already holds the workbook" {
        New-Item -ItemType Directory -Force (Join-Path $global:Base "out\release") | Out-Null
        Set-Content (Join-Path $global:Base "out\release\jira2dashboard.xlsm") "wb"
        Invoke-Script "manual-test.ps1" | Out-Null
        (Calls) -contains "run build" | Should Be $false
    }
    It "-Workbook and -Port are passed on, and the workbook is not built" {
        $r = Invoke-Script "manual-test.ps1" -Named @{ Workbook = "C:\Downloads\x.xlsm"; Port = 8100 }
        $r.Code | Should Be 0
        (Calls)[0] | Should Be "stop-mock -Port 8100"
        (Calls)[1] | Should Be "run mock-workbook --workbook C:\Downloads\x.xlsm --port 8100"
        (Calls)[2] | Should Be "start-mock -Port 8100"
    }
    It "-Python runs the extraction on the copy and opens the xlsx, no macro" {
        $r = Invoke-Script "manual-test.ps1" -Named @{ Python = $true }
        $r.Code | Should Be 0
        (Calls)[-1] | Should Match "^run extract --workbook .*jira2dashboard4manualtest.xlsm$"
        ($r.Out -replace "\s+", " ") | Should Match "opens in a new Excel: the data, the results and the charts, no macro"
        Assert-MockCalled Start-Process -Scope It -Times 1 -ParameterFilter { ($ArgumentList -join " ") -match "jira2dashboard.xlsx" }
    }
    It "stops when the build fails" {
        $global:BuildExit = 1
        $r = Invoke-Script "manual-test.ps1"
        $r.Code | Should Be 1
        $r.Out | Should Match "the build of the workbook failed"
    }
    It "stops when the copy cannot be prepared" {
        $global:MockWbExit = 1
        (Invoke-Script "manual-test.ps1" -Named @{ Workbook = "x.xlsm" }).Out | Should Match "the copy could not be prepared"
    }
    It "stops when the fake Jira does not start" {
        $global:StartExit = 1
        (Invoke-Script "manual-test.ps1" -Named @{ Workbook = "x.xlsm" }).Out | Should Match "the fake Jira did not start"
        Assert-MockCalled Start-Process -Scope It -Times 0
    }
    It "-Python stops when the extraction fails" {
        $global:ExtractExit = 1
        (Invoke-Script "manual-test.ps1" -Named @{ Workbook = "x.xlsm"; Python = $true }).Out | Should Match "the Python extraction failed"
        Assert-MockCalled Start-Process -Scope It -Times 0
    }
    It "writes its log, and goes on when the log cannot be written" {
        Invoke-Script "manual-test.ps1" -Named @{ Workbook = "x.xlsm" } | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "manual-test") | Should Be $true
        Block-Logs
        (Invoke-Script "manual-test.ps1" -Named @{ Workbook = "x.xlsm" }).Code | Should Be 0
    }
}

# --- activate.ps1 -------------------------------------------------------------------------------------------------------

Describe "activate.ps1" {
    BeforeEach { Reset-Sandbox; Set-Fake "chcp" }
    AfterEach { Restore-Sandbox }

    It "asks for setup.ps1 when there is no .venv" {
        $ErrorActionPreference = "Continue"
        $text = (. (Join-Path $global:Scripts "activate.ps1") *>&1 | Out-String)
        $text | Should Match "No .venv yet: run scripts.setup.ps1 first"
        (Calls).Count | Should Be 0
    }
    It "activates .venv, sets UTF-8 and PYTHONPATH to src, in the session that dot-sources it" {
        New-Item -ItemType Directory -Force (Join-Path $global:Base ".venv\Scripts") | Out-Null
        Set-Content (Join-Path $global:Base ".venv\Scripts\Activate.ps1") '$env:J2D_ACTIVATED = "yes"'
        $text = (. (Join-Path $global:Scripts "activate.ps1") *>&1 | Out-String)
        $text | Should Match "jira2dashboard venv active"
        $env:J2D_ACTIVATED | Should Be "yes"
        $env:PYTHONIOENCODING | Should Be "utf-8"
        $env:PYTHONPATH | Should Be (Join-Path $global:Base "src")
        (Calls) | Should Be @("chcp 65001")
    }
}

# --- scenario.ps1 -------------------------------------------------------------------------------------------------------

Describe "scenario.ps1" {
    BeforeEach {
        Reset-Sandbox
        $global:StopExit = 0; $global:BuildExit = 0; $global:MockWbExit = 0; $global:StartExit = 0; $global:ExtractExit = 0
        $global:ExcelExit = 0; $global:WriteXlsx = "ok"
        # the scripts it plays are replaced by functions named like their full path: this test is about the scenario itself
        Set-Fake "stop-mock" { $global:LASTEXITCODE = $global:StopExit } -name (Join-Path $global:Scripts "stop-mock.ps1")
        Set-Fake "start-mock" { $global:LASTEXITCODE = $global:StartExit } -name (Join-Path $global:Scripts "start-mock.ps1")
        Set-Fake "run" {
            if ($args[0] -eq "build") { New-Item -ItemType Directory -Force "out\release" | Out-Null; Set-Content "out\release\jira2dashboard.xlsm" "wb"; $global:LASTEXITCODE = $global:BuildExit }
            if ($args[0] -eq "mock-workbook") { $global:LASTEXITCODE = $global:MockWbExit }
            if ($args[0] -eq "extract") {
                New-Item -ItemType Directory -Force "out\manual_test" | Out-Null
                if ($global:WriteXlsx -eq "ok") { Set-Content "out\manual_test\jira2dashboard.xlsx" "xlsx" }
                if ($global:WriteXlsx -eq "empty") { New-Item "out\manual_test\jira2dashboard.xlsx" -ItemType File | Out-Null }
                $global:LASTEXITCODE = $global:ExtractExit
            }
            if ($args[0] -eq "test-excel") { $global:LASTEXITCODE = $global:ExcelExit }
        } -name (Join-Path $global:Scripts "run.ps1")
    }
    AfterEach { Restore-Sandbox }

    It "plays the steps in order and stops the fake Jira at the end" {
        $r = Invoke-Script "scenario.ps1"
        "$($r.Code)|$($r.Out)" | Should Match "^0\|"
        $r.Out | Should Match "all steps passed"
        (Calls | ForEach-Object { ($_ -split " ")[0..1] -join " " }) -join "," | Should Be "stop-mock -Port,run build,run mock-workbook,start-mock -Port,run extract,stop-mock -Port"
        (Calls)[3] | Should Be "start-mock -Port 8099 -Hidden"
        (Calls)[4] | Should Match "^run extract --workbook .*jira2dashboard4manualtest.xlsm$"
        $r.Out | Should Match "step 1: stop an earlier fake Jira"
        $r.Out | Should Match "(?s)scenario: summary.*7 *stop the fake Jira *ok"
    }
    It "does not build when the workbook exists" {
        New-Item -ItemType Directory -Force (Join-Path $global:Base "out\release") | Out-Null
        Set-Content (Join-Path $global:Base "out\release\jira2dashboard.xlsm") "wb"
        $r = Invoke-Script "scenario.ps1"
        $r.Out | Should Match "exists, not built"
        (Calls) -contains "run build" | Should Be $false
    }
    It "-Workbook and -Port are passed on" {
        Set-Content (Join-Path $global:Base "given.xlsm") "wb"
        $r = Invoke-Script "scenario.ps1" -Named @{ Workbook = (Join-Path $global:Base "given.xlsm"); Port = 8100 }
        "$($r.Code)|$($r.Out)" | Should Match "^0\|"
        (Calls)[0] | Should Be "stop-mock -Port 8100"
        (Calls)[1] | Should Match "^run mock-workbook --workbook .*given.xlsm --port 8100$"
        (Calls)[2] | Should Be "start-mock -Port 8100 -Hidden"
    }
    It "fails when the workbook given does not exist, and still stops the fake Jira" {
        $r = Invoke-Script "scenario.ps1" -Named @{ Workbook = (Join-Path $global:Base "missing.xlsm") }
        "$($r.Code)|$($r.Out)" | Should Match "^1\|"
        ($r.Out -replace "\s+", " ") | Should Match "FAILED \(J2D-SCR-041 - the workbook .*missing.xlsm does not exist\)"
        $r.Out | Should Match "skipped \(an earlier step failed\)"
        (Calls)[-1] | Should Be "stop-mock -Port 8099"
    }
    It "fails and skips the next steps when the build fails" {
        $global:BuildExit = 1
        $r = Invoke-Script "scenario.ps1"
        "$($r.Code)|$($r.Out)" | Should Match "^1\|"
        $r.Out | Should Match "step 2: FAILED \(exit code 1\)"
        (Calls) -contains "start-mock -Port 8099 -Hidden" | Should Be $false
        (Calls)[-1] | Should Be "stop-mock -Port 8099"
    }
    It "fails when the copy cannot be prepared" {
        $global:MockWbExit = 1
        (Invoke-Script "scenario.ps1").Code | Should Be 1
    }
    It "fails when the fake Jira does not start" {
        $global:StartExit = 1
        $r = Invoke-Script "scenario.ps1"
        "$($r.Code)|$($r.Out)" | Should Match "^1\|"
        (Calls) -contains "run extract --workbook $(Join-Path $global:Base 'out\manual_test\jira2dashboard4manualtest.xlsm')" | Should Be $false
    }
    It "fails when the extraction fails, and still stops the fake Jira" {
        $global:ExtractExit = 1
        $r = Invoke-Script "scenario.ps1"
        "$($r.Code)|$($r.Out)" | Should Match "^1\|"
        $r.Out | Should Match "FAILED"
        $r.Out | Should Match "scenario: J2D-SCR-040 - FAILED. Log: "
        (Calls)[-1] | Should Be "stop-mock -Port 8099"
    }
    It "fails when the xlsx was not written" {
        $global:WriteXlsx = "none"
        ((Invoke-Script "scenario.ps1").Out -replace "\s+", " ") | Should Match "FAILED \(J2D-SCR-042 - .*jira2dashboard.xlsx was not written\)"
    }
    It "fails when the xlsx is empty" {
        $global:WriteXlsx = "empty"
        ((Invoke-Script "scenario.ps1").Out -replace "\s+", " ") | Should Match "J2D-SCR-043 - .*is empty"
    }
    It "-Excel plays the Excel scenarios too, and a failure of them fails the scenario" {
        $r = Invoke-Script "scenario.ps1" -Named @{ Excel = $true }
        "$($r.Code)|$($r.Out)" | Should Match "^0\|"
        (Calls) -contains "run test-excel" | Should Be $true
        $global:FakeCalls.Clear()
        $global:ExcelExit = 1
        (Invoke-Script "scenario.ps1" -Named @{ Excel = $true }).Code | Should Be 1
    }
    It "-KeepMock leaves the fake Jira running and says how to stop it" {
        $r = Invoke-Script "scenario.ps1" -Named @{ KeepMock = $true }
        "$($r.Code)|$($r.Out)" | Should Match "^0\|"
        @(Calls | Where-Object { $_ -like "stop-mock*" }).Count | Should Be 1               # only the one of the start
        $r.Out | Should Match "the fake Jira still runs on port 8099: stop it with .\\scripts\\stop-mock.ps1 -Port 8099"
    }
    It "writes its log, and goes on when the log cannot be written" {
        Invoke-Script "scenario.ps1" | Out-Null
        Stop-Transcript | Out-Null
        Test-Path (Log-File "scenario") | Should Be $true
        Block-Logs
        (Invoke-Script "scenario.ps1").Code | Should Be 0
    }
}

# --- one transcript when a script calls another -------------------------------------------------------------------------------

Describe "a script called by another one" {
    BeforeEach { Reset-Sandbox }
    AfterEach { Restore-Sandbox }

    It "stop-mock.ps1 starts no transcript when its caller already logs" {
        $J2D_LogOwner = "caller"
        $r = Invoke-Script "stop-mock.ps1"
        $r.Code | Should Be 0
        $r.Out | Should Match "no fake Jira started"
        Test-Path (Log-File "stop-mock") | Should Be $false
    }
    It "run.ps1 starts no transcript when its caller already logs" {
        New-Venv
        $J2D_LogOwner = "caller"
        (Invoke-Script "run.ps1" -Positional @("test")).Code | Should Be 0
        Test-Path (Log-File "run_test") | Should Be $false
    }
    It "scenario.ps1 keeps one log with the lines of the scripts it plays, and no log of theirs" {
        Set-Fake "start-mock" { $global:LASTEXITCODE = 0 } -name (Join-Path $global:Scripts "start-mock.ps1")
        Set-Fake "run" {
            if ($args[0] -eq "extract") { New-Item -ItemType Directory -Force "out\manual_test" | Out-Null; Set-Content "out\manual_test\jira2dashboard.xlsx" "xlsx" }
            if ($args[0] -eq "build") { New-Item -ItemType Directory -Force "out\release" | Out-Null; Set-Content "out\release\jira2dashboard.xlsm" "wb" }
            $global:LASTEXITCODE = 0
        } -name (Join-Path $global:Scripts "run.ps1")
        $r = Invoke-Script "scenario.ps1"                                     # the real stop-mock.ps1 runs inside it
        $r.Code | Should Be 0
        Stop-Transcript | Out-Null
        (Get-Content (Log-File "scenario") -Raw) | Should Match "stop-mock: no fake Jira started"
        (Get-Content (Log-File "scenario") -Raw) | Should Not Match "Impossible de d|Start-Transcript"
        Test-Path (Log-File "stop-mock") | Should Be $false
    }
}
