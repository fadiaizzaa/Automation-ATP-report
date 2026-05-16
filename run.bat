@echo off
setlocal
cd /d "%~dp0"

set "ROOT=%CD%"
set "PYTHONPATH=%ROOT%\pipeline\run;%ROOT%\pipeline\cpq;%ROOT%\pipeline"
set "CONFIG=%ROOT%\config\ingest.yml"
set "VENV=%ROOT%\.venv\Scripts\activate.bat"

if not exist "%CONFIG%" (
  echo Config not found: %CONFIG%
  echo Copy config\ingest.example.yml to config\ingest.yml and edit it first.
  pause
  exit /b 1
)

if not exist "%VENV%" (
  echo Virtual environment not found. Run once:
  echo   python -m venv .venv
  echo   .venv\Scripts\activate
  echo   pip install -r requirements.txt
  pause
  exit /b 1
)

call "%VENV%"

echo.
echo Running week pipeline ...
echo Config: %CONFIG%
echo.

python "%ROOT%\pipeline\run\run_week_pipeline.py" --config "%CONFIG%"
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Pipeline finished OK. Check your week folder under ATP NMS for updated workbooks.
) else (
  echo Pipeline FAILED ^(exit %EXITCODE%^). Read the messages above or send the log to support.
)
echo.
pause
exit /b %EXITCODE%
