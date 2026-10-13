@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "ROOT=%~dp0.."
pushd "%ROOT%"
title Document pipeline - Quark DOCX first

set "INPUT="
if exist "input\word\*.docx" set "INPUT=input\word"
if exist "input\word\*.doc" set "INPUT=input\word"
if not defined INPUT set "INPUT=input\pdfs"

if not exist "output\word" mkdir "output\word"
echo.
echo Input: %ROOT%\%INPUT%
echo Output: %ROOT%\output\word
echo Mode: Quark DOCX first; OOXML compatibility cleanup; no AI/model inference
echo.
python -u scripts\quark_pdf_to_word.py --input "%INPUT%" --output "output\word" --pdf-engine auto %*
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" goto complete
echo Failed. Check output\reports and tmp\word.
goto finish

:complete
echo Complete. Reports are in output\reports.

:finish
pause
popd
exit /b %RC%
