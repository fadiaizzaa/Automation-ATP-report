@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
  echo Create the venv first: python -m venv .venv
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"

echo.
echo Repairing pywin32 for Excel / xlwings ...
pip install --force-reinstall pywin32
python ".venv\Scripts\pywin32_postinstall.py" -install

echo.
echo Testing Excel automation ...
python -c "import pywintypes, win32com.client; import xlwings as xw; print('engine:', xw.engines.active); app=xw.App(visible=False, add_book=False); app.quit(); print('Excel automation OK')"

if "%ERRORLEVEL%"=="0" (
  echo.
  echo Setup complete. You can run run.bat again.
) else (
  echo.
  echo Setup failed. Ensure Microsoft Excel is installed and open it once manually.
  echo You can still run prepare with: python pipeline\run\prepare_week_inputs.py --config config\ingest.yml --skip-excel-recalc
)
echo.
pause
exit /b %ERRORLEVEL%
