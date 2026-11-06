@echo off
rem Launcher of scripts\release.ps1: double-click it, or run it from a console. release.cmd --help explains it.
setlocal EnableExtensions
set "J2D_SCRIPT=release"
set "J2D_LAUNCHER=%~f0"
set "J2D_WHAT=Builds a release in out\release\VERSION; -Publish creates the GitHub release (it publishes: use it with care)."
set "J2D_NEEDS_ARGS=1"
set "J2D_EXAMPLE=release.cmd -Version v1.0.0"
"%~dp0_launch.cmd" %*
exit /b %errorlevel%
