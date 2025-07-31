# FDA 510(k) Intelligence Suite - Deployment Guide

This guide provides comprehensive instructions for deploying the FDA 510(k) Intelligence Suite to various environments, from individual workstations to enterprise-wide deployments.

## 📋 Table of Contents
- [Prerequisites](#prerequisites)
- [Build Process](#build-process)
- [Deployment Methods](#deployment-methods)
- [Configuration Management](#configuration-management)
- [Environment-Specific Deployments](#environment-specific-deployments)
- [Troubleshooting](#troubleshooting)
- [Maintenance](#maintenance)

## Prerequisites

### System Requirements
- **Operating System**: Windows 10+ (x64)
- **Microsoft Excel**: 2016+ with VBA and Power Query support
- **PowerShell**: 5.1+ (for automated deployment)
- **Network Access**: Internet connection for FDA API and optional OpenAI integration
- **User Permissions**: Local admin rights for initial setup (optional for ongoing use)

### Required Excel References
The following VBA references must be available:
- Microsoft Scripting Runtime
- Microsoft Visual Basic for Applications Extensibility 5.3
- Microsoft XML, v6.0

### Network Requirements
- **Outbound HTTPS**: Port 443 to `api.fda.gov` (required)
- **Outbound HTTPS**: Port 443 to `api.openai.com` (optional, for OpenAI features)
- **Proxy Settings**: Configure Excel/Windows proxy settings if behind corporate firewall

## Build Process

### 1. Pre-Build Preparation
```bash
# Clone the repository
git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git
cd FDA-510k-Intelligence-Suite

# Verify source files
ls assets/excel-workbooks/
ls src/vba/
ls src/powerquery/
```

### 2. Build Commands

#### Standard Production Build
```bash
# Using npm scripts (recommended)
npm run build

# Or directly with PowerShell
.\scripts\build\build.ps1

# Or with batch file
scripts\build\build.bat
```

#### Development Build (with debug modules)
```bash
npm run build:debug
# Or
.\scripts\build\build.ps1 -IncludeDebug
```

### 3. Build Artifacts
After successful build, the `dist/` folder contains:
- `FDA_510k_Intelligence_Suite_v[VERSION].xlsm` - Main application file
- `build_info.json` - Build metadata
- `README.txt` - End-user installation instructions

## Deployment Methods

### Method 1: Single Workstation Deployment

#### Manual Deployment
1. **Prepare Target Location**
   ```powershell
   # Create deployment folder
   New-Item -ItemType Directory -Path "C:\Applications\FDA510k" -Force
   ```

2. **Copy Files**
   - Copy built `.xlsm` file to target location
   - Copy `README.txt` for user reference

3. **Configure User Environment**
   - Create OpenAI key directory: `%APPDATA%\510k_Tool\`
   - Set up desktop shortcut (optional)

#### Automated Single Deployment
```bash
# Deploy to local machine
.\scripts\deploy\deploy.ps1 -DeploymentPath "C:\Applications\FDA510k"

# Deploy with backup of existing version
.\scripts\deploy\deploy.ps1 -DeploymentPath "C:\Applications\FDA510k" -CreateBackup:$true
```

### Method 2: Network Share Deployment

#### 1. Prepare Network Location
```powershell
# Create network share (run on file server)
New-Item -ItemType Directory -Path "\\fileserver\shares\Applications\FDA510k" -Force
```

#### 2. Deploy to Network Share
```bash
# Deploy to network location
.\scripts\deploy\deploy.ps1 -DeploymentPath "\\fileserver\shares\Applications\FDA510k"
```

#### 3. User Access Setup
Users can either:
- **Option A**: Run directly from network share (requires good network performance)
- **Option B**: Copy to local machine and run locally (recommended)

### Method 3: Enterprise Group Policy Deployment

#### 1. Create MSI Package (Optional)
For enterprise environments, consider packaging the Excel file in an MSI:
```powershell
# Use tools like Advanced Installer or WiX Toolset
# Package the .xlsm file with proper registry entries
```

#### 2. Group Policy Deployment
1. **Prepare SYSVOL Share**
   ```
   \\domain.com\SYSVOL\domain.com\Policies\{PolicyGUID}\Machine\Applications\
   ```

2. **Configure Software Installation Policy**
   - Computer Configuration > Policies > Software Settings > Software Installation
   - Add new package pointing to MSI or direct .xlsm file

#### 3. User Configuration
- Deploy OpenAI key setup script via Group Policy
- Configure Excel Trust Center settings via Group Policy

### Method 4: SCCM/Intune Deployment

#### SCCM Application Deployment
1. **Create Application**
   - Detection method: File existence check for the .xlsm file
   - Installation command: Copy operation or MSI installation
   - Uninstall command: File deletion script

2. **Requirements**
   - Microsoft Excel 2016+ installed
   - Windows 10+ operating system

#### Intune Win32 App Deployment
1. **Prepare IntuneWin Package**
   ```powershell
   # Use Microsoft Win32 Content Prep Tool
   IntuneWinAppUtil.exe -c "source_folder" -s "FDA_510k_Intelligence_Suite.xlsm" -o "output_folder"
   ```

2. **Configure Deployment**
   - Install command: PowerShell script to copy and configure
   - Detection rules: File or registry key
   - Requirements: Excel 2016+, Windows 10+

## Configuration Management

### Pre-Deployment Configuration

#### 1. Customize mod_Config.bas
```vba
' Update these constants before building
Public Const MAINTAINER_USERNAME As String = "YourDomainAdmin"
Public Const DEFAULT_AC_WEIGHT As Double = 0.2
' ... other settings as needed
```

#### 2. Environment-Specific Builds
```bash
# Build for different environments
.\scripts\build\build.ps1 -Version "1.0-DEV"   # Development
.\scripts\build\build.ps1 -Version "1.0-PROD"  # Production
```

### Post-Deployment Configuration

#### 1. OpenAI API Key Setup
For users requiring OpenAI features:
```powershell
# Automated setup script
$apiKeyPath = "$env:APPDATA\510k_Tool"
New-Item -ItemType Directory -Path $apiKeyPath -Force
# Note: Actual API key must be provided by user or admin
```

#### 2. Excel Trust Center Configuration
```powershell
# Registry settings for Trust Center (requires admin rights)
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
Set-ItemProperty -Path $regPath -Name "VBAWarnings" -Value 1  # Enable all macros
```

## Environment-Specific Deployments

### Development Environment
- Enable debug logging
- Use test data sources
- Include all debugging modules
- Shorter refresh intervals for testing

### Staging Environment
- Production-like configuration
- Limited user access
- Full logging enabled
- Performance monitoring

### Production Environment
- Optimized for performance
- Minimal logging
- Production API endpoints
- User training materials included

## Troubleshooting

### Common Deployment Issues

#### 1. Excel Security Warnings
**Problem**: Users see security warnings and macros don't run
**Solution**:
```powershell
# Configure Trust Center via Group Policy or registry
# Or instruct users to enable macros manually
```

#### 2. Network Connectivity Issues
**Problem**: FDA API requests fail
**Solution**:
- Verify internet connectivity
- Check corporate firewall settings
- Configure proxy settings in Excel

#### 3. VBA Reference Issues
**Problem**: "Missing reference" errors
**Solution**:
- Verify required Office components are installed
- Re-register VBA references if needed
- Use Excel Repair if necessary

#### 4. OpenAI Integration Failures
**Problem**: Company recaps show "Needs Research"
**Solution**:
- Verify API key is correctly placed and formatted
- Check OpenAI API quota and billing
- Validate network access to api.openai.com

### Diagnostic Commands
```powershell
# Check Excel version
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\*\Excel\InstallRoot" | Select PSChildName, Path

# Verify network connectivity
Test-NetConnection -ComputerName "api.fda.gov" -Port 443
Test-NetConnection -ComputerName "api.openai.com" -Port 443

# Check PowerShell version
$PSVersionTable.PSVersion
```

## Maintenance

### Regular Maintenance Tasks

#### 1. Monthly Updates
- Check for new versions of the application
- Review and update scoring weights if needed
- Monitor API usage and quotas

#### 2. Quarterly Reviews
- Review user feedback and feature requests
- Update documentation
- Performance optimization

#### 3. Annual Tasks
- Security review and compliance check
- API key rotation (OpenAI)
- User training refreshers

### Version Updates

#### 1. Prepare Update
```bash
# Build new version
npm run build

# Test in staging environment
# Notify users of upcoming update
```

#### 2. Deploy Update
```bash
# Backup current version
.\scripts\deploy\deploy.ps1 -DeploymentPath "\\server\share" -CreateBackup:$true -OverwriteConfig:$false

# Users can update by running the new version
```

#### 3. Post-Update Validation
- Verify all users can access the new version
- Check logs for any errors
- Gather user feedback

### Monitoring and Logging

#### Application Logs
- Location: Excel RunLog sheet
- Contains: API calls, errors, performance metrics
- Review regularly for issues

#### Deployment Logs
- Location: `deployment_config.json` in deployment folder
- Contains: Deployment history, versions, timestamps
- Use for tracking rollouts

### Backup and Recovery

#### Backup Strategy
```powershell
# Regular backup of deployment
$backupPath = "\\server\backups\FDA510k\$(Get-Date -Format 'yyyy-MM-dd')"
Copy-Item -Path "\\server\deploy\FDA510k" -Destination $backupPath -Recurse
```

#### Recovery Procedures
1. **Application Recovery**: Restore from backup deployment
2. **Data Recovery**: Archive sheets contain historical data
3. **Configuration Recovery**: Use config JSON files as reference

## Security Considerations

### Data Protection
- FDA data is public but should be handled according to corporate data policies
- OpenAI API keys must be protected as sensitive credentials
- Consider data retention policies for archived information

### Access Control
- Limit deployment/admin access to authorized personnel
- Consider read-only deployments for most users
- Implement proper file permissions on network shares

### Compliance
- Ensure deployment meets corporate IT policies
- Document security configurations
- Regular security reviews and updates

## Support and Documentation

### User Support
- Provide `README.txt` with each deployment
- Maintain internal wiki or knowledge base
- Establish support contact information

### Administrator Resources
- Keep deployment scripts and documentation current
- Maintain configuration templates
- Document customizations and environment-specific settings

---

For additional support or questions about deployment, please refer to the project documentation or contact the development team.