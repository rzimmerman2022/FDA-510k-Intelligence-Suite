# FDA 510(k) Intelligence Suite

**Professional-grade FDA 510(k) clearance data analysis tool built in Excel VBA with Power Query integration.**

Automated fetching, scoring, caching, and archiving of FDA 510(k) clearance data directly within Microsoft Excel. This enterprise-ready tool streamlines the process of identifying and prioritizing relevant FDA 510(k) clearances based on configurable scoring models.

## 🚀 Quick Start

1. **Clone the Repository**
   ```bash
   git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git
   cd FDA-510k-Intelligence-Suite
   ```

2. **Open the Main Workbook**
   - Navigate to `assets/excel-workbooks/`
   - Open the main `.xlsm` file
   - Enable macros and content when prompted

3. **Configure Your Settings**
   - Set your username in `src/vba/utilities/mod_Config.bas`
   - Configure scoring weights in the Excel `Weights` sheet
   - (Optional) Set up OpenAI API for company summaries

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

### **VBA Module Organization**
```
src/vba/
├── core/                         # Business Logic
│   ├── mod_510k_Processor.bas    # Main processing orchestration
│   ├── mod_Archive.bas           # Archive management
│   ├── mod_Cache.bas             # Company data caching
│   ├── mod_Schema.bas            # Data schema management
│   ├── mod_Score.bas             # Scoring algorithm
│   └── mod_Weights.bas           # Weight configuration
├── utilities/                    # Shared Utilities
│   ├── mod_Config.bas            # Global configuration
│   ├── mod_DataIO.bas            # Data input/output operations
│   ├── mod_Format.bas            # Formatting and UI
│   ├── mod_Utils.bas             # General utilities
│   └── mod_*Debug*.bas           # Debug and logging utilities
└── modules/                      # Application Modules
    ├── mod_RefreshSolutions.bas  # Power Query refresh solutions
    ├── ModuleManager.bas         # VBA code management
    └── ThisWorkbook.cls          # Workbook event handlers
```

### **Running Tests**
```bash
# Unit tests
cd tests/unit
# Open test files in Excel VBA editor

# Integration tests  
cd tests/integration
# Follow test documentation in docs/technical-specs/
```

### **Build Process**
```bash
# Development build
cd scripts/build
# Run build scripts as documented

# Production deployment
cd scripts/deploy  
# Follow deployment guides in docs/implementation-guides/
```

## 📊 System Requirements

- **Microsoft Excel**: 2016+ with Power Query and VBA support
- **Windows OS**: Required for Windows-specific dependencies
- **Internet Connection**: For openFDA API and optional OpenAI calls
- **VBA References**: 
  - Microsoft Scripting Runtime
  - Microsoft Visual Basic for Applications Extensibility 5.3
  - Microsoft XML, v6.0

## 🔐 Configuration

### **Required Setup**
1. **Maintainer Username**: Update `MAINTAINER_USERNAME` in `src/vba/utilities/mod_Config.bas`
2. **Scoring Parameters**: Configure tables in Excel `Weights` sheet
3. **API Endpoints**: Verify openFDA API configuration

### **Optional Features**
- **OpenAI Integration**: Set API key for automated company summaries
- **Advanced Logging**: Configure debug modes and trace levels
- **Custom Formatting**: Adjust UI constants and formatting rules

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