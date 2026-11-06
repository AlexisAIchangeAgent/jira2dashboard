@echo off
rem Launcher of scripts\start-mock.ps1: double-click it, or run it from a console. start-mock.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=start-mock"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Starts the fake Jira on http://127.0.0.1:8099 (never the real one) and prints the python.exe that serves it."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
