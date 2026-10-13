@echo off
setlocal EnableExtensions
pushd "%~dp0.."
python "%~dp0run_mechanical_two_pages.py"
popd
pause
