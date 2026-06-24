@echo off
echo Building risk_report_generator.exe ...
echo.
.venv\Scripts\pyinstaller --onefile --name risk_report_generator --collect-all questionary launcher.py

echo.
echo Copying to OneDrive...
set "ONEDRIVE_DIR=D:\HII\OneDrive\HydroDataSci\Project\province_risk_report_automation"
copy /Y dist\risk_report_generator.exe "%ONEDRIVE_DIR%\"
copy /Y config.yaml "%ONEDRIVE_DIR%\"
xcopy /E /I /Y template "%ONEDRIVE_DIR%\template"
mkdir "%ONEDRIVE_DIR%\output" 2>nul

echo.
echo Done! risk_report_generator.exe is ready in OneDrive.
pause
