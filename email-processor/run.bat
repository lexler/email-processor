@echo off

set "TARGET_DIR=%~1"
if "%TARGET_DIR%"=="" set "TARGET_DIR=C:\Users\wilsora3\Documents\PCIT Automated"

echo Running email processor...
uv run process_emails.py "%TARGET_DIR%"