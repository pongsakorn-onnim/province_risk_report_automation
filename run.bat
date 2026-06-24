@echo off
set /p PROVINCE="Province (e.g. สุโขทัย): "
set /p REGION="Region (e.g. เหนือ): "
set /p DATA_DIR="Data dir (e.g. D:\HII\OneDrive\HydroDataSci\Data\Risk_Area\Risk_Forecast\summary\202606_summary): "
set /p YYYYMM="Period YYYYMM (e.g. 202606): "

echo.
echo Running report...
.venv\Scripts\python main.py --province "%PROVINCE%" --region "%REGION%" --data-dir "%DATA_DIR%" --yyyymm %YYYYMM%

echo.
pause
