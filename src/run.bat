@echo off
setlocal
cd /d "%~dp0"

py ".\TalkAlert.py"
if errorlevel 1 (
  python ".\TalkAlert.py"
)

pause