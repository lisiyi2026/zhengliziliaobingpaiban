@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "ROOT=%~dp0.."
pushd "%ROOT%"
title Document pipeline
set "PYTHON=python"
if exist "D:\python\python3.12\python.exe" set "PYTHON=D:\python\python3.12\python.exe"
set "PYTHONUTF8=1"

if not exist "input\word" mkdir "input\word"
if not exist "input\pdfs" mkdir "input\pdfs"
if not exist "output\word" mkdir "output\word"

rem Normal processing is Word-first and does not require any bridge.
rem LibreOffice is only needed for .doc conversion or explicit rendering.
echo Word-first: Quark DOCX -> formatting -> OOXML compatibility cleanup.
echo No AI/model inference is used.

echo.
echo Input: %ROOT%\input\word
echo Reference PDFs: %ROOT%\input\pdfs
echo Output: %ROOT%\output\word
echo.
"%PYTHON%" -u scripts\quark_pdf_to_word.py --input "input\word" --output "output\word" --bridge %*
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (echo Complete. Reports are in output\reports.) else (echo Failed. Check output\reports or tmp\lo-bridge\errors-v2.)
popd
pause
exit /b %RC%
