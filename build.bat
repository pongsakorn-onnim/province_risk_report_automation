@echo off
echo Building run_report.exe ...
echo.
.venv\Scripts\pyinstaller --onefile --name run_report --collect-all questionary launcher.py
echo.
echo Done! Copy  dist\run_report.exe  into the OneDrive project folder.
pause
