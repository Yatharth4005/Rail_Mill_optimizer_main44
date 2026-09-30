@echo off
title Rail Mill Optimizer Server (172.17.0.20:5674)
echo ===================================================
echo   Jindal Steel - Rail Mill Optimizer Server
echo ===================================================
echo.
echo Target Server Address: http://172.17.0.20:5674
echo.
echo [1/3] Running Database Migrations...
python manage.py migrate
echo.
echo [2/3] Importing Master Calculation Rules...
python manage.py import_rail_mill_length_calc
echo.
echo [3/3] Starting Rail Mill Optimizer Server on 0.0.0.0:5674...
echo Access in browser: http://http://172.17.0.20:5674 or http://localhost:5674
echo Press Ctrl+C to stop the server.
echo ===================================================
echo.
python manage.py runserver 0.0.0.0:5674
pause
