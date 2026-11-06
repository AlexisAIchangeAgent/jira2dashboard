@echo off
rem Internal helper of the launchers scripts\<name>.cmd: not meant to be run by hand.
rem The launcher sets J2D_SCRIPT (the .ps1 without its extension), J2D_WHAT (one line), J2D_LAUNCHER (its own full path),
rem and optionally J2D_NEEDS_ARGS=1 (a double-click only explains) with J2D_EXAMPLE, or J2D_CONSOLE=1 (a console stays open).
setlocal EnableExtensions EnableDelayedExpansion

rem 1. Wait for a key at the end only when the window would close by itself (a double-click gives cmd.exe /c ""<launcher>" "),
rem    or when J2D_PAUSE=1.
set "pause_at_end="
if "%J2D_PAUSE%"=="1" set "pause_at_end=1"
set "seen=%cmdcmdline%"
set "expected=%comspec% /c ""%J2D_LAUNCHER%" ""
if /i "!seen!"=="!expected!" set "pause_at_end=1"

rem 2. Help: this launcher, then the help of the script.
for %%h in (--help -help -h) do if /i "%~1"=="%%h" goto :help
if "%~1"=="/?" goto :help

rem 3. A script that cannot do anything without arguments is not run by a double-click.
if not defined J2D_NEEDS_ARGS goto :run
if not "%~1"=="" goto :run
call :usage
echo.
echo Not run: this script needs arguments. Open a console in this folder and give them, for example:
echo   %J2D_EXAMPLE%
set "code=2"
goto :done

:run
if defined J2D_CONSOLE goto :console
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0%J2D_SCRIPT%.ps1" %*
set "code=!errorlevel!"
goto :done

:console
rem activate.ps1 is meant to be dot-sourced: a console of its own, with the environment activated, stays open.
powershell -NoExit -NoProfile -ExecutionPolicy Bypass -Command ". '%~dp0%J2D_SCRIPT%.ps1'"
set "code=!errorlevel!"
goto :done

:help
call :usage
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0%J2D_SCRIPT%.ps1" -Help
set "code=!errorlevel!"
goto :done

:done
if defined pause_at_end (
    echo.
    echo Exit code: !code!
    echo Press any key to close this window...
    pause >nul
)
exit /b !code!

:usage
echo %J2D_SCRIPT%.cmd - launcher of scripts\%J2D_SCRIPT%.ps1
echo %J2D_WHAT%
echo.
echo Usage:
echo   Double-click               runs it with its defaults, then waits for a key so that you can read the output.
echo   %J2D_SCRIPT%.cmd [arguments]   from a console: the arguments go to %J2D_SCRIPT%.ps1 unchanged, no wait.
echo   %J2D_SCRIPT%.cmd --help        this text, then the help of %J2D_SCRIPT%.ps1 with its arguments.
echo.
echo It runs: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\%J2D_SCRIPT%.ps1
echo The execution policy is bypassed for this run only: nothing is changed on the machine, no administrator rights are needed.
echo Set J2D_PAUSE=1 to wait for a key at the end also from a console.
echo.
exit /b 0
