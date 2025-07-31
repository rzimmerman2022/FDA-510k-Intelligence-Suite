# FDA 510(k) Intelligence Suite

**Professional-grade FDA 510(k) clearance data analysis tool built in Excel VBA with Power Query integration.**

Automated fetching, scoring, caching, and archiving of FDA 510(k) clearance data directly within Microsoft Excel. This enterprise-ready tool streamlines the process of identifying and prioritizing relevant FDA 510(k) clearances based on configurable scoring models.

## 🚀 Quick Start

### Prerequisites
- Microsoft Excel 2016+ with VBA and Power Query support
- Windows 10+ operating system
- Internet connection for FDA API access

### Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git
   cd FDA-510k-Intelligence-Suite
   ```

2. **Build the Application**
   ```bash
   # Windows Command Prompt
   scripts\build\build.bat
   
   # Or PowerShell
   .\scripts\build\build.ps1
   ```

3. **Open the Built Workbook**
   - Navigate to `dist/` folder (created by build)
   - Open the versioned `.xlsm` file
   - Enable macros and content when prompted

4. **Configure Your Settings**
   - The build process automatically configures your username
   - Configure scoring weights in the Excel `Weights` sheet
   - (Optional) Set up OpenAI API key for company summaries

### Optional: OpenAI Integration
For automated company summaries, add your API key:
1. Open: `%APPDATA%\510k_Tool\openai_key.txt`
2. Replace placeholder with your actual OpenAI API key
3. Get API key from: https://platform.openai.com/api-keys

## 📁 Repository Structure

```
FDA-510k-Intelligence-Suite/
├── .archives/                    # Historical files and diagnostics
│   ├── commits/                  # Commit messages and history
│   └── diagnostics/              # Diagnostic and troubleshooting files
├── .github/                      # GitHub configuration
│   ├── ISSUE_TEMPLATE/           # Issue templates
│   └── workflows/                # CI/CD workflows
├── assets/                       # Binary assets and workbooks
│   └── excel-workbooks/          # Main Excel files (.xlsm)
├── config/                       # Configuration files
├── docs/                         # Documentation
│   ├── images/                   # Documentation images
│   ├── implementation-guides/    # Implementation and setup guides
│   └── technical-specs/          # Technical specifications
├── samples/                      # Sample data and outputs
├── scripts/                      # Build and deployment scripts
│   ├── build/                    # Build scripts
│   └── deploy/                   # Deployment scripts
├── src/                          # Source code
│   ├── powerquery/               # Power Query (.pq) files
│   └── vba/                      # VBA source code
│       ├── _legacy/              # Legacy code (archived)
│       ├── core/                 # Core business logic modules
│       ├── modules/              # Application-specific modules
│       └── utilities/            # Utility and helper modules
└── tests/                        # Test files
    ├── integration/              # Integration tests
    └── unit/                     # Unit tests
```

## 🔧 Core Features

### **Automated Data Pipeline**
- **Power Query Integration**: Dynamically pulls FDA 510(k) data from openFDA API
- **Smart Refresh Logic**: Fetches previous month's data automatically
- **Connection Management**: Robust error handling and connection cleanup

### **Advanced Scoring Engine**
- **Multi-Factor Scoring**: Advisory Committee, Product Code, Keywords, Submission Type, Processing Time, Geography
- **Configurable Weights**: Manage scoring parameters via Excel tables
- **Negative Factors**: Built-in rules for cosmetic, diagnostic classifications
- **Synergy Bonuses**: Keyword combinations and specific criteria bonuses

### **Enterprise Features**
- **Company Intelligence**: Local caching with optional OpenAI-powered summaries
- **Automated Archiving**: Monthly data archiving with static value conversion
- **Comprehensive Logging**: Multi-level logging system with performance optimization
- **Maintainer Controls**: Role-based feature access and bypass capabilities

### **Professional UI/UX**
- **Dynamic Formatting**: Score-based conditional formatting and styling
- **Smart Column Management**: Automated width adjustment and duplicate handling
- **Device Name Optimization**: Truncation with hover-to-view full text
- **Modular Architecture**: Clean separation of concerns across VBA modules

## 🛠️ Development Workflow

### **Project Structure**
```
FDA-510k-Intelligence-Suite/
├── assets/excel-workbooks/       # Source Excel files
├── config/                       # Configuration files (JSON)
├── docs/                         # Documentation
├── scripts/                      # Build and deployment automation
│   ├── build/                    # Build scripts (PowerShell + Batch)
│   └── deploy/                   # Deployment scripts
├── src/                          # Source code
│   ├── powerquery/               # Power Query (.pq) files  
│   └── vba/                      # VBA modules
│       ├── core/                 # Business logic
│       ├── utilities/            # Shared utilities
│       └── modules/              # Application modules
└── tests/                        # Test files
```

### **Build Commands**
```bash
# Build for production
npm run build

# Build with debug modules
npm run build:debug

# Clean build artifacts
npm run clean
```

### **Deployment**
```bash
# Deploy to specific location
scripts\deploy\deploy.bat "C:\Path\To\Deployment"

# Deploy with PowerShell (more options)
.\scripts\deploy\deploy.ps1 -DeploymentPath "C:\Path" -CreateBackup:$true
```

### **Development Setup**
1. Make changes to VBA files in `src/vba/`
2. Update Power Query in `src/powerquery/`
3. Build using `npm run build`
4. Test the built workbook from `dist/` folder
5. Deploy to target environment

### **Testing**
- Unit tests: Located in `tests/unit/`
- Integration tests: Located in `tests/integration/`  
- Manual testing: Use built workbook in `dist/` folder

## 📊 System Requirements

- **Microsoft Excel**: 2016+ with Power Query and VBA support
- **Windows OS**: Required for Windows-specific dependencies
- **Internet Connection**: For openFDA API and optional OpenAI calls
- **VBA References**: 
  - Microsoft Scripting Runtime
  - Microsoft Visual Basic for Applications Extensibility 5.3
  - Microsoft XML, v6.0

## 🔐 Configuration

### **Automated Configuration**
The build process automatically configures:
- ✅ Maintainer username (from system)
- ✅ API key directory structure  
- ✅ Default scoring parameters
- ✅ Logging and debug settings

### **Manual Configuration**
Configure these settings in Excel after building:

1. **Scoring Weights** (Required)
   - Open the built workbook
   - Navigate to `Weights` sheet
   - Update scoring tables as needed

2. **OpenAI API Key** (Optional)
   - File location: `%APPDATA%\510k_Tool\openai_key.txt`
   - Required only for automated company summaries
   - Get key from: https://platform.openai.com/api-keys

3. **Advanced Settings** (Optional)
   - Debug modes: Set in VBA `mod_Config.bas`
   - Custom formatting: Adjust constants in modules
   - Logging levels: Configure in application settings

### **Configuration Files**
- `config/app.config.json` - Application settings reference
- `config/environment.json` - Environment-specific settings
- `package.json` - Project metadata and scripts

## 📖 Documentation

- **[Architecture Guide](docs/ARCHITECTURE.md)**: System design and component relationships
- **[Development Guide](docs/AI_DEVELOPMENT_GUIDE.md)**: Coding standards and AI assistance guidelines  
- **[User Guide](docs/USER_GUIDE.md)**: End-user operation instructions
- **[Implementation Guides](docs/implementation-guides/)**: Setup and deployment instructions
- **[Technical Specifications](docs/technical-specs/)**: Detailed technical documentation

## 🤝 Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests.

## 📄 License

This project is licensed under the terms specified in [LICENSE](LICENSE).

## 🆘 Support

For issues, questions, or contributions:
- Create an issue in this repository
- Review documentation in the `docs/` directory
- Check implementation guides for common setup issues

---

**Built with ❤️ for FDA regulatory intelligence and compliance teams**