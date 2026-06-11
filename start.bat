@echo off
echo ================================================
echo  The Awareness Group - CMO Weekly Report
echo ================================================
echo.

if "%ANTHROPIC_API_KEY%"=="" (
  set /p ANTHROPIC_API_KEY="Paste your Anthropic API key (for Executive Summary AI): "
)

echo.
echo Starting server at http://localhost:5151
echo Open that URL in your browser.
echo.
python app.py
pause
