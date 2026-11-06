@echo off
rem Launcher of scripts\setup.ps1: double-click it, or run it from a console. setup.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=setup"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Creates .venv (Python 3.14) and installs the dependencies; -Dev adds the test tools."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
