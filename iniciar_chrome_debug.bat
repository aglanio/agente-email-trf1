@echo off
echo Iniciando Chrome com porta de debug para automacao PJe...
echo.
echo IMPORTANTE: Feche o Chrome atual primeiro!
echo.
taskkill /f /im chrome.exe 2>nul
timeout /t 2 /nobreak >nul
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data"
echo Chrome iniciado com debug port 9222
echo Agora pode usar "Puxar Arquivos do PJe"
pause
