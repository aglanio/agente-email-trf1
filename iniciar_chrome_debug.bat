@echo off
echo === Agente Email TRF1 - Chrome Debug ===
echo.
echo Fechando Chrome...
taskkill /f /im chrome.exe 2>nul
timeout /t 3 /nobreak >nul
echo Iniciando Chrome com debug port 9222...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%USERPROFILE%\ChromeDebug" ^
  --remote-allow-origins=* ^
  --no-first-run ^
  http://localhost:8090
timeout /t 5 /nobreak >nul
echo.
echo Chrome iniciado com debug port 9222!
echo O certificado digital do Windows funciona normalmente.
echo.
echo Agora no Agente Email, clique "Puxar Tudo PJe"
echo.
pause
