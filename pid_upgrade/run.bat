@echo off
title AVEVA P&ID Upgrade Script
echo ========================================
echo AVEVA P&ID Upgrade Script
echo ========================================
echo.
echo Starting upgrade process...
echo.

REM Run PowerShell script with bypass policy
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0pid-upgrade.ps1"

echo.
echo ========================================
echo Script execution completed.
echo Check D:\upgrade_log.txt for details.
echo ========================================
echo.
pause