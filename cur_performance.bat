@echo off
setlocal
cd /d "%~dp0"

set "ROOT=%CD%"
set "PYTHONPATH=%ROOT%\pipeline\run;%ROOT%\pipeline\cpq;%ROOT%\pipeline"
set "CONFIG=%ROOT%\config\ingest.yml"
set "VENV=%ROOT%\.venv\Scripts\activate.bat"

if not exist "%CONFIG%" (
  echo Config not found: %CONFIG%
  pause
  exit /b 1
)

if not exist "%VENV%" (
  echo Virtual environment not found. Run setup from README.md first.
  pause
  exit /b 1
)

call "%VENV%"

echo.
echo Daily Current Performance update ...
echo Drop ZIP files in: performance_daily under your week raw files folder
echo Config: %CONFIG%
echo.

python "%ROOT%\pipeline\cur_performance.py" --config "%CONFIG%"
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Current Performance update finished OK.
) else (
  echo Update FAILED ^(exit %EXITCODE%^).
)
echo.
pause
exit /b %EXITCODE%
