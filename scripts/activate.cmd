@echo off
rem Launcher of scripts\activate.ps1: double-click it, or run it from a console. activate.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=activate"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Opens a PowerShell console with the environment of .venv activated (the script is meant to be dot-sourced)."
set "J2D_CONSOLE=1"
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
