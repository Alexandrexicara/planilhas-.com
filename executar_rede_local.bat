@echo off
REM ====================================================
REM  PLANILHAS.COM - MODO REDE LOCAL (WiFi compartilhado)
REM  Permite acesso de outros PCs na mesma rede WiFi
REM ====================================================
echo.
echo ========================================================
echo  INICIANDO SERVIDOR EM MODO REDE LOCAL
echo  Outros PCs na mesma rede WiFi poderao acessar
echo ========================================================
echo.

REM Liberar porta no firewall (apenas primeira vez - ignora erro se ja existe)
netsh advfirewall firewall add rule name="Planilhas LAN 5000" dir=in action=allow protocol=TCP localport=5000 >nul 2>&1

REM Ativar modo LAN
set PLANILHAS_LAN=1
set PORT=5000

REM Iniciar o servidor
python app.py

pause
