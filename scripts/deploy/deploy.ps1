# FDA 510(k) Intelligence Suite - Deployment Script
# This script deploys the Excel workbook to specified locations

param(
    [Parameter(Mandatory=$true)]
    [string]$DeploymentPath,
    
    [Parameter(Mandatory=$false)]
    [string]$SourcePath = "..\..\dist",
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateBackup = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$OverwriteConfig = $false
)

Write-Host "FDA 510(k) Intelligence Suite - Deployment Process" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# Validate deployment path
if (!(Test-Path $DeploymentPath)) {
    Write-Host "ERROR: Deployment path does not exist: $DeploymentPath" -ForegroundColor Red
    exit 1
}

# Find the latest build
$fullSourcePath = Join-Path $PSScriptRoot $SourcePath
if (!(Test-Path $fullSourcePath)) {
    Write-Host "ERROR: Source path does not exist: $fullSourcePath" -ForegroundColor Red
    Write-Host "Please run the build script first." -ForegroundColor Yellow
    exit 1
}

$latestBuild = Get-ChildItem -Path $fullSourcePath -Filter "FDA_510k_Intelligence_Suite_v*.xlsm" |
               Sort-Object LastWriteTime -Descending |
               Select-Object -First 1

if (!$latestBuild) {
    Write-Host "ERROR: No build found in: $fullSourcePath" -ForegroundColor Red
    Write-Host "Please run the build script first." -ForegroundColor Yellow
    exit 1
}

Write-Host "Deploying: $($latestBuild.Name)" -ForegroundColor Yellow
Write-Host "To: $DeploymentPath" -ForegroundColor Yellow

# Create backup if requested
if ($CreateBackup) {
    $existingFile = Get-ChildItem -Path $DeploymentPath -Filter "*.xlsm" |
                    Where-Object { $_.Name -like "FDA_510k_Intelligence_Suite*" } |
                    Select-Object -First 1
    
    if ($existingFile) {
        $backupName = "$($existingFile.BaseName)_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsm"
        $backupPath = Join-Path $DeploymentPath $backupName
        
        Copy-Item -Path $existingFile.FullName -Destination $backupPath -Force
        Write-Host "Created backup: $backupName" -ForegroundColor Green
    }
}

# Deploy the workbook
$deploymentFile = Join-Path $DeploymentPath $latestBuild.Name
Copy-Item -Path $latestBuild.FullName -Destination $deploymentFile -Force
Write-Host "Deployed workbook successfully" -ForegroundColor Green

# Deploy additional files
$buildInfo = Join-Path $fullSourcePath "build_info.json"
if (Test-Path $buildInfo) {
    Copy-Item -Path $buildInfo -Destination $DeploymentPath -Force
    Write-Host "Deployed build info" -ForegroundColor Green
}

$readme = Join-Path $fullSourcePath "README.txt"
if (Test-Path $readme) {
    Copy-Item -Path $readme -Destination $DeploymentPath -Force
    Write-Host "Deployed README" -ForegroundColor Green
}

# Create/Update deployment config
$deployConfig = @{
    DeploymentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    DeployedFile = $latestBuild.Name
    DeployedBy = $env:USERNAME
    DeploymentMachine = $env:COMPUTERNAME
    BackupCreated = $CreateBackup.IsPresent
}

$configPath = Join-Path $DeploymentPath "deployment_config.json"
if ((Test-Path $configPath) -and !$OverwriteConfig) {
    Write-Host "Deployment config exists (use -OverwriteConfig to replace)" -ForegroundColor Yellow
} else {
    $deployConfig | ConvertTo-Json | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "Created deployment config" -ForegroundColor Green
}

# Create quick access batch file
$batchContent = @"
@echo off
echo Starting FDA 510(k) Intelligence Suite...
start "" "$($latestBuild.Name)"
"@

$batchPath = Join-Path $DeploymentPath "Start_FDA_510k_Tool.bat"
$batchContent | Out-File -FilePath $batchPath -Encoding ASCII
Write-Host "Created quick start batch file" -ForegroundColor Green

Write-Host "`nDeployment completed successfully!" -ForegroundColor Green
Write-Host "Users can run 'Start_FDA_510k_Tool.bat' to open the application" -ForegroundColor Cyan