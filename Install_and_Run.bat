@echo off
echo ===================================================
echo   Auto Certificate Generator - Initializing...
echo ===================================================
echo.
echo Installing requirements... (This might take a minute on the first run)
python -m pip install -r requirements.txt
echo.
echo Launching the application...
python AutoCert.py
pause
