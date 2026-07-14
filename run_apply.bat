@echo off
setlocal
cd /d "%~dp0"
python file_automation.py --config config.example.json --apply --confirm REPLACE
pause
