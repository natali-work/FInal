@echo off
echo Processing Excel files...
python process_excel.py
echo.
echo Running analysis...
python analyze_experiment_v3.py
echo.
echo Done!
pause

