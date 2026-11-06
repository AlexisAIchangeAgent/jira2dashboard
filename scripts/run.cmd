@echo off
rem Launcher of scripts\run.ps1: double-click it, or run it from a console. run.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=run"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Runs one command of the project: test, coverage, build, diagnose... A double-click lists the commands."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
