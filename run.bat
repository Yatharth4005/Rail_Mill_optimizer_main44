@echo off
setlocal
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  py -m venv .venv
)
call .venv\Scripts\activate.bat
py -m pip install --upgrade pip
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
py manage.py migrate
start "Rail Mill Optimizer" http://127.0.0.1:8000/
py manage.py runserver 127.0.0.1:8000
