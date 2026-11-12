@echo off
rem Launcher of scripts\bootstrap.ps1: double-click it, or run it from a console. bootstrap.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=bootstrap"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Installs, checks the machine, tests and builds the .xlsm and the .xlsx in one go."
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
