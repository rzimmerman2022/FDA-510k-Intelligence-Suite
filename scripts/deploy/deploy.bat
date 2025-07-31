@echo off
REM FDA 510(k) Intelligence Suite - Deployment Script (Windows Batch)
REM Wrapper for PowerShell deployment script

echo FDA 510(k) Intelligence Suite - Deployment Process
echo =================================================

REM Check if PowerShell is available
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: PowerShell is not available on this system
    exit /b 1
)

REM Check if deployment path is provided
if "%~1"=="" (
    echo ERROR: Please provide a deployment path
    echo Usage: deploy.bat "C:\Path\To\Deployment\Folder"
    echo.
    echo Optional parameters:
    echo   -CreateBackup:$false     Skip creating backup
    echo   -OverwriteConfig:$true   Overwrite existing config
    pause
    exit /b 1
)

REM Run the PowerShell deployment script
powershell -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" -DeploymentPath %*

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Deployment failed. See above for details.
    pause
    exit /b %errorlevel%
)

echo.
echo Deployment completed successfully!
pause