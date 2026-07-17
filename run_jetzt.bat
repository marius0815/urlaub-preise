@echo off
REM Manueller Trigger: Doppelklick startet die Flugpreis-Abfrage sofort.
REM Argumente werden durchgereicht, z.B.:  run_jetzt.bat --reise Kenia
cd /d "%~dp0"
echo Starte Flugpreis-Abfrage ...
python flugpreise.py %*
echo.
echo Fertig. Ergebnisse: best-of.xlsx / lauf.log
pause
