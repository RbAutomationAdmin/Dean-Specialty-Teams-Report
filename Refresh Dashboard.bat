@echo off
rem One-click refresh: rebuilds specialty_teams_dashboard.html from today's Qlik exports in Downloads.
cd /d "%~dp0"
python refresh_dashboard.py %*
echo.
pause
