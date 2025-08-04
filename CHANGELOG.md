```
 ██████╗██╗  ██╗ █████╗ ███╗   ██╗ ██████╗ ███████╗██╗      ██████╗  ██████╗ 
██╔════╝██║  ██║██╔══██╗████╗  ██║██╔════╝ ██╔════╝██║     ██╔═══██╗██╔════╝ 
██║     ███████║███████║██╔██╗ ██║██║  ███╗█████╗  ██║     ██║   ██║██║  ███╗
██║     ██╔══██║██╔══██║██║╚██╗██║██║   ██║██╔══╝  ██║     ██║   ██║██║   ██║
╚██████╗██║  ██║██║  ██║██║ ╚████║╚██████╔╝███████╗███████╗╚██████╔╝╚██████╔╝
 ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝ ╚═════╝ ╚══════╝╚══════╝ ╚═════╝  ╚═════╝ 
```

# Changelog

> All notable changes to the FDA 510(k) Intelligence Suite will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased] - 2025-08-04

### 🎨 Documentation Overhaul
- **Enhanced README.md**: 
  - Added ASCII art header for visual appeal
  - Comprehensive AI-friendly documentation with code patterns
  - Detailed pipeline architecture with mermaid diagrams
  - Module dependency graph for better understanding
  - Complete data schema reference table
  - Testing checklist for contributors
- **Created PIPELINE_SUMMARY.md**: Standalone pipeline documentation
- **Updated CHANGELOG.md**: Added ASCII header and better formatting
- **Improved ARCHITECTURE.md**: Added data flow diagram and clearer structure

### 🧹 Repository Cleanup
- **Moved to .archives**:
  - `_Analysis/` (empty folder)
  - `_Archive/` (empty folder) 
  - `_Documents/` (15 working files)
  - `_Reports/` (1 Excel file)
  - `_Research/` (methodology documents)
- **Removed**:
  - `src/vba/_legacy (do not use!)/` folder with 12 legacy VBA files
- **Result**: Clean, professional repository structure

### 📚 AI Assistant Enhancements
- Added comprehensive coding context for AI assistants
- Included VBA-specific patterns and constraints
- Documented common tasks with code examples
- Added performance optimization guidelines
- Included memory management best practices

### Added
- Comprehensive build and deployment automation system
- Configuration management with JSON reference files
- Professional-grade project structure following industry standards
- npm-style package.json with script commands
- PowerShell and batch build scripts for cross-compatibility
- Automated deployment with backup and configuration management
- OpenAI API key directory structure with setup instructions

### Changed
- Complete repository restructure to gold standard practices
- Updated README.md with modern build/deploy workflow instructions
- Enhanced CONTRIBUTING.md with comprehensive development guidelines
- Improved configuration documentation and automation
- Maintainer username automatically configured from system

### Fixed
- Missing build pipeline infrastructure
- Empty configuration and script directories
- Manual configuration requirements streamlined
- API key setup process automated and documented

### Infrastructure
- Added package.json for project metadata and scripts
- Created config/ directory with JSON reference files
- Implemented scripts/build/ with PowerShell and batch automation
- Added scripts/deploy/ with deployment automation
- Established proper directory structure for enterprise deployment

## [1.9.0] - 2025-04-30

### Added
- Split Code Generation architecture
- Enhanced VBA module organization with core/utilities/modules separation
- Comprehensive error handling and logging system
- Company caching with OpenAI integration
- Automated archiving to monthly sheets
- Power Query refresh solutions and diagnostics

### Changed
- Modular VBA architecture for better maintainability
- Improved scoring algorithm with configurable weights
- Enhanced data processing pipeline with error recovery
- Better separation of concerns across modules

### Fixed
- Power Query connection management issues
- Synchronous refresh problems resolved
- Column mapping and data schema improvements
- Memory management and performance optimizations

## [1.0.0] - Initial Release

### Added
- Core FDA 510(k) data processing functionality
- Power Query integration with openFDA API
- Basic scoring algorithm implementation
- Excel VBA automation framework
- Data archiving capabilities
- Company information caching

### Features
- Automated monthly FDA 510(k) data retrieval
- Multi-factor scoring system (AC, PC, ST, Keywords, Geography)
- Excel table-based weight configuration
- Conditional formatting and data visualization
- Error logging and debugging capabilities

---

## Types of Changes
- **Added** for new features
- **Changed** for changes in existing functionality  
- **Deprecated** for soon-to-be removed features
- **Removed** for now removed features
- **Fixed** for any bug fixes
- **Security** for vulnerability fixes
- **Infrastructure** for build/deployment/tooling changes