@echo off
title Excel Image Sheet & SQL Server
echo =========================================================================
echo    EXCEL IMAGE SPREADSHEET & LOCAL SQL DATABASE
echo =========================================================================
echo.
echo  [+] Dang khoi dong Web Server va Ket noi CSDL SQL...
echo  [+] Ung dung Local: http://localhost:8000
echo  [+] Ung dung GitHub Pages: https://khanghohy.github.io/proexcel/
echo  [+] Duong dan CSDL SQL: %~dp0data\excel_app.db
echo.
if exist "%~dp0cloudflared.exe" (
  echo  [+] Dang khoi dong Cloudflare Tunnel ket noi GitHub Pages...
  start /b "" "%~dp0cloudflared.exe" tunnel --url http://localhost:8000
)
echo =========================================================================
echo.

cd /d "%~dp0"
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
pause
