@echo off
setlocal EnableExtensions
chcp 65001 >nul
pushd "%~dp0.."
title LibreOffice bridge
set "PYTHON=python"
if exist "D:\python\python3.12\python.exe" set "PYTHON=D:\python\python3.12\python.exe"
"%PYTHON%" "%~dp0lo_user_bridge.py"
popd
pause
