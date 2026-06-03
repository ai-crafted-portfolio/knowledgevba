@echo off
setlocal
cd /d "C:\kvba\push"
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\kvba\push\_fix_encoding.ps1"
echo EXIT=%ERRORLEVEL%
endlocal
