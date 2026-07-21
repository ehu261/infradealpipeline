@echo off
cd /d "%~dp0"
python update_market_intel.py %*
pause
