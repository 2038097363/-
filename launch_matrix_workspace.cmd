@echo off
cd /d "%~dp0"
if exist "%~dp0MatrixWorkspace.exe" (
    start "" "%~dp0MatrixWorkspace.exe"
    exit /b 0
)
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0matrix_calculator.ps1"
