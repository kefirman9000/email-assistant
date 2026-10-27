@echo off
title Email Assistant setup
cd /d "%~dp0"
echo.
echo  Setting up Email Assistant. This only needs to be done once.
echo.

where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo  Python is not installed.
  echo  Please install it from https://www.python.org/downloads/
  echo  and tick "Add python.exe to PATH" during installation.
  echo.
  pause
  exit /b 1
)

echo  [1/2] Installing the parts the app needs...
%PY% -m pip install --upgrade -r requirements.txt
if errorlevel 1 (
  echo.
  echo  Something went wrong installing. Please send a photo of this window
  echo  to whoever set this up for you.
  pause
  exit /b 1
)

echo.
echo  [2/2] Downloading the AI model. The first time can take 10+ minutes...
where ollama >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Ollama is not installed. Please install it from https://ollama.com/download
  echo  then run this Setup again.
  pause
  exit /b 1
)
ollama pull llama3.1:8b

echo.
echo  All done! Double-click "Start Email Assistant" to begin.
echo.
pause
