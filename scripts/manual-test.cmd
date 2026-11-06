@echo off
rem Launcher of scripts\manual-test.ps1: double-click it, or run it from a console. manual-test.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=manual-test"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Fills a copy of the workbook for the fake Jira, starts the fake Jira and opens the copy in a new Excel."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
