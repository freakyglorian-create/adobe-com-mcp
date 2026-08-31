@echo off
REM Adobe COM MCP Server launcher
REM Requires: Python 3.10+ with pywin32 + mcp installed (pip install pywin32 "mcp<2")

set "PYTHON=python"
set "WORKDIR=%~dp0"

cd /d "%WORKDIR%"
"%PYTHON%" server.py
