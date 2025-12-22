@echo off
setlocal
cd /d "%~dp0"

git pull

if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv
  call ".venv\Scripts\activate.bat"
  python -m pip install -r requirements.txt
) else (
  call ".venv\Scripts\activate.bat"
)

start "" /b ".venv\Scripts\pythonw.exe" tkinter_etiqueta.py
endlocal
