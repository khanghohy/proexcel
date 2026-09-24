@echo off
title Excel Image Sheet & SQL Server
echo =========================================================================
echo    EXCEL IMAGE SPREADSHEET & LOCAL SQL DATABASE
echo =========================================================================
echo.
echo  [+] Dang khoi dong Web Server va Ket noi CSDL SQL...
echo  [+] Ung dung: http://localhost:8000
echo  [+] Duong dan CSDL: %~dp0data\excel_app.db
echo.
echo  Nhan Ctrl + C trong cua so nay de dung server bat cu luc nao.
echo =========================================================================
echo.

cd /d "%~dp0"
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
pause
