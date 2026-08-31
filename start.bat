@echo off
REM Launch the Adobe Photoshop COM MCP server (stdio).
REM Use this as the MCP "command" in TRAE settings, or configure node + server.js directly.
cd /d "%~dp0"
node server.js
