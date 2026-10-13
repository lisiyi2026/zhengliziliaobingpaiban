@echo off
setlocal
chcp 65001 >nul
pushd "%~dp0.."
title PDF-Extract-Kit formula bridge
set "PYTHON=%~dp0..\.venv-pdfkit\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
"%PYTHON%" "%~dp0pdfkit_user_bridge.py"
popd
pause
