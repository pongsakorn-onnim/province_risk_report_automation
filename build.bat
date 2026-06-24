@echo off
echo Building run_report.exe ...
echo.
.venv\Scripts\pyinstaller --onefile --name risk_report_generator --collect-all questionary launcher.py
echo.
echo Done! Copy  dist\risk_report_generator.exe  into the OneDrive project folder.
pause
