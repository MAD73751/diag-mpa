@echo off
cd /d "%~dp0"
echo ============================================================
echo   Outils MaPrimeAdapt - DMAAI
echo ============================================================

if not exist "logos_b64.py" (echo ERREUR: logos_b64.py manquant & pause & exit /b 1)
if not exist "agent_diagnostic_mpa.html" (echo ERREUR: agent_diagnostic_mpa.html manquant & pause & exit /b 1)
if not exist "serveur_diagnostic_mpa.py" (echo ERREUR: serveur_diagnostic_mpa.py manquant & pause & exit /b 1)
if not exist "regles_anah_2026.json" (echo ERREUR: regles_anah_2026.json manquant - referentiel des regles Anah & pause & exit /b 1)

echo Arret ancien serveur sur port 5052...
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr :5052 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
timeout /t 1 /nobreak >nul

echo Installation dependances...
pip install flask flask-cors python-docx anthropic -q 2>nul

echo.
echo ============================================================
echo   DIAGNOSTIC  : http://localhost:5052
echo   RECEPTION   : http://localhost:5052/reception
echo   CTRL+C pour arreter le serveur
echo ============================================================
timeout /t 2 /nobreak >nul

REM Ouvrir les deux onglets dans Firefox
start "" "http://localhost:5052"
timeout /t 1 /nobreak >nul
start "" "http://localhost:5052/reception"

python serveur_diagnostic_mpa.py
pause
