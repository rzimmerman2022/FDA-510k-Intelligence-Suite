@echo off
REM FDA 510(k) Intelligence Suite - Build Script (Windows Batch)
REM Simple wrapper for PowerShell build script

echo FDA 510(k) Intelligence Suite - Build Process
echo =============================================

REM Check if PowerShell is available
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: PowerShell is not available on this system
    exit /b 1
)

REM Run the PowerShell build script
powershell -ExecutionPolicy Bypass -File "%~dp0build.ps1" %*

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed. See above for details.
    pause
    exit /b %errorlevel%
)

echo.
echo Build completed successfully!
pause