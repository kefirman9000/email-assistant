@echo off
cd /d "%~dp0"
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw "%~dp0email_assistant.py"
) else (
  start "" pythonw "%~dp0email_assistant.py"
)
