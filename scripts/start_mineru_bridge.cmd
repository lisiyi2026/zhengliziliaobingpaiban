@echo off
setlocal EnableExtensions
chcp 65001 >nul
pushd "%~dp0.."
title MinerU bridge
set "PYTHON=%~dp0..\.venv-mineru\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
"%PYTHON%" "%~dp0mineru_user_bridge.py"
popd
pause
