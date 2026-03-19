@echo off
title Agente E-mail TRF1 - Local
echo ==========================================
echo   Agente E-mail TRF1 - Modo Local
echo   Outlook sera aberto automaticamente!
echo ==========================================
echo.

cd /d "%~dp0"

set PYTHON=C:\Users\aglan\AppData\Local\Programs\Python\Python312\python.exe

echo Iniciando servidor na porta 8090...
echo Acesse: http://localhost:8090
echo.
"%PYTHON%" -m uvicorn app:app --host 0.0.0.0 --port 8090 --reload

pause
