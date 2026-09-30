@echo off
title Rail Mill Optimizer Server - Production (172.17.18.13:5674)
echo ===================================================
echo   Jindal Steel - Rail Mill Optimizer Server (Production)
echo ===================================================
echo.
echo Target Server Address: http://172.17.18.13:5674
echo.
echo [1/3] Running Database Migrations...
py manage.py migrate
echo.
echo [2/3] Importing Master Calculation Rules...
py manage.py import_rail_mill_length_calc
echo.
echo [3/3] Starting Waitress WSGI Server on 0.0.0.0:5674...
echo Access in browser: http://172.17.18.13:5674 or http://localhost:5674
echo ===================================================
echo.
py -m waitress --host=0.0.0.0 --port=5674 rail_casting_plan.wsgi:application
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Waitress WSGI server not installed or exited with error.
    echo Falling back to Django server on 0.0.0.0:5674...
    py manage.py runserver 0.0.0.0:5674
)
pause
