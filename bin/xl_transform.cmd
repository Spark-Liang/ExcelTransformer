@echo off
set PROJECT_ROOT=%~dp0\..\
set PYTHONPATH=%PROJECT_ROOT%\src
set PYTHON_RT=%PROJECT_ROOT%\venv\Scripts\python.exe

%PYTHON_RT% -m xl_transform.main.App %*