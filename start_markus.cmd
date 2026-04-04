@echo off
setlocal
:: Silent launcher — uses pyw.exe so no console window ever appears.
:: If using a venv, point to its pyw.exe instead:
::   "%~dp0venv\Scripts\pyw.exe" "%~dp0markus.py"
"C:\WINDOWS\pyw.exe" "%~dp0markus.py"
endlocal
