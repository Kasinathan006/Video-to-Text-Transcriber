@echo off
REM VoxDoc AI — production server launcher
cd /d "%~dp0"
echo Starting VoxDoc AI Server...
python api_server.py
pause
