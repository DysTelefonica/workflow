@echo off
echo === access-sandbox sync ===
echo.
echo Estableciendo password...
setx ACCESS_SANDBOX_PW "dpddpd"
set ACCESS_SANDBOX_PW=dpddpd
echo.
echo Ejecutando sync-backends.ps1...
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\sync-backends.ps1" -ConfigPath "%~dp0configs\backends_config.json"
echo.
pause
