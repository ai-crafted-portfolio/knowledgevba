@echo off
rem knowledgevba v2.3 one-shot installer launcher (ASCII only)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0knowledgevba_install.ps1"
exit /b %errorlevel%
