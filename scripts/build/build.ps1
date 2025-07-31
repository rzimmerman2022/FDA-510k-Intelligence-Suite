# FDA 510(k) Intelligence Suite - Build Script
# This script prepares the Excel workbook for distribution

param(
    [Parameter(Mandatory=$false)]
    [string]$Version = "1.0",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "..\..\dist",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDebug = $false
)

Write-Host "FDA 510(k) Intelligence Suite - Build Process" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# Create output directory if it doesn't exist
$fullOutputPath = Join-Path $PSScriptRoot $OutputPath
if (!(Test-Path $fullOutputPath)) {
    New-Item -ItemType Directory -Path $fullOutputPath -Force | Out-Null
    Write-Host "Created output directory: $fullOutputPath" -ForegroundColor Green
}

# Define source and destination paths
$sourceWorkbook = Join-Path $PSScriptRoot "..\..\assets\excel-workbooks"
$latestWorkbook = Get-ChildItem -Path $sourceWorkbook -Filter "*.xlsm" | 
                  Where-Object { $_.Name -notlike "*backup*" -and $_.Name -notlike "*old*" } |
                  Sort-Object LastWriteTime -Descending | 
                  Select-Object -First 1

if (!$latestWorkbook) {
    Write-Host "ERROR: No Excel workbook found in assets/excel-workbooks" -ForegroundColor Red
    exit 1
}

Write-Host "Building from: $($latestWorkbook.Name)" -ForegroundColor Yellow

# Copy workbook to output
$outputFileName = "FDA_510k_Intelligence_Suite_v$Version.xlsm"
$outputFile = Join-Path $fullOutputPath $outputFileName
Copy-Item -Path $latestWorkbook.FullName -Destination $outputFile -Force
Write-Host "Copied workbook to: $outputFileName" -ForegroundColor Green

# Create build info file
$buildInfo = @{
    Version = $Version
    BuildDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    SourceFile = $latestWorkbook.Name
    BuildMachine = $env:COMPUTERNAME
    BuildUser = $env:USERNAME
    IncludesDebug = $IncludeDebug.IsPresent
}

$buildInfoPath = Join-Path $fullOutputPath "build_info.json"
$buildInfo | ConvertTo-Json | Out-File -FilePath $buildInfoPath -Encoding UTF8
Write-Host "Created build info file" -ForegroundColor Green

# Create README for distribution
$readmeContent = @"
# FDA 510(k) Intelligence Suite v$Version

## Installation Instructions

1. **Enable Macros**: Open the Excel file and enable macros when prompted
2. **Trust Access**: Go to File > Options > Trust Center > Trust Center Settings > Macro Settings
   - Check "Trust access to the VBA project object model"
3. **Configure Settings**: Update the configuration in the VBA editor:
   - Open VBA Editor (Alt+F11)
   - Navigate to mod_Config
   - Update MAINTAINER_USERNAME to your Windows username
4. **Set Up API Key** (Optional - for OpenAI features):
   - Create file at: %APPDATA%\510k_Tool\openai_key.txt
   - Add your OpenAI API key to the file

## First Run

1. The tool will automatically fetch the previous month's FDA 510(k) data
2. Check the RunLog sheet for any errors
3. Review the CurrentMonthData sheet for results

## Support

For issues or questions, please refer to the documentation or contact your administrator.

Build Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@

$readmePath = Join-Path $fullOutputPath "README.txt"
$readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
Write-Host "Created distribution README" -ForegroundColor Green

# Clean up any temporary files
if (!$IncludeDebug) {
    # Remove debug modules if not including debug
    Write-Host "Note: Debug modules retained in build (remove manually if needed)" -ForegroundColor Yellow
}

Write-Host "`nBuild completed successfully!" -ForegroundColor Green
Write-Host "Output location: $fullOutputPath" -ForegroundColor Cyan