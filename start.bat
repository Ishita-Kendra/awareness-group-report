@echo off
echo ================================================
echo  The Awareness Group - CMO Weekly Report
echo ================================================
echo.

set ENV_FILE=%~dp0.env

if not exist "%ENV_FILE%" (
  echo First-time setup: your Anthropic API key is needed for AI features.
  echo It will be saved in a local .env file and never asked again.
  echo.
  set /p API_KEY="Paste your Anthropic API key (sk-ant-...): "
  echo ANTHROPIC_API_KEY=%API_KEY%> "%ENV_FILE%"
  echo.
  echo Key saved to .env
  echo.
)

echo Starting server at http://localhost:5151
echo Open that address in your browser.
echo Press Ctrl+C to stop.
echo.
cd /d "%~dp0"
python app.py
pause
