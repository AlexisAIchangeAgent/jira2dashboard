@echo off
rem Launcher of scripts\scenario.ps1: double-click it, or run it from a console. scenario.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=scenario"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Plays the whole scenario against the fake Jira: filled copy, Python extraction, checks, clean-up. Logs each step."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
