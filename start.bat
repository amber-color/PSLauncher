@echo off
cd /d %~dp0
start "" powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "PSLauncher.ps1"
exit