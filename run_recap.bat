@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_recap.ps1" %*
endlocal
