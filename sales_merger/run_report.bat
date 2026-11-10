@echo off
rem One-click run: put sales files into the "input" folder and double-click this file.
cd /d "%~dp0"
if not exist .venv (
    echo First run: setting up Python environment...
    python -m venv .venv || goto :error
    .venv\Scripts\python -m pip install -q -r requirements.txt || goto :error
)
.venv\Scripts\python merge_sales.py || goto :error
start "" output
pause
exit /b 0

:error
echo.
echo Something went wrong. See the message above.
pause
exit /b 1
