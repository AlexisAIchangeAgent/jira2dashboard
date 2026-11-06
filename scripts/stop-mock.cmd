@echo off
rem Launcher of scripts\stop-mock.ps1: double-click it, or run it from a console. stop-mock.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=stop-mock"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Stops the fake Jira that start-mock started, and nothing else."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
