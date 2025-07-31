# Contributing to FDA 510(k) Intelligence Suite

We welcome contributions to the FDA 510(k) Intelligence Suite! This document provides comprehensive guidelines for contributing to this enterprise-grade medical device regulatory intelligence tool.

## =Ë Table of Contents
- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Workflow](#development-workflow)
- [Code Standards](#code-standards)
- [Testing Guidelines](#testing-guidelines)
- [Documentation Standards](#documentation-standards)
- [Submission Process](#submission-process)

## Code of Conduct

This project and everyone participating in it is governed by our Code of Conduct. By participating, you are expected to uphold this code. Report unacceptable behavior to the project maintainers.

## Getting Started

### Prerequisites
- Windows 10+ operating system
- Microsoft Excel 2016+ with VBA and Power Query
- Git for version control
- PowerShell 5.1+ for build scripts

### Development Environment Setup

1. **Clone and Build**
   ```bash
   git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git
   cd FDA-510k-Intelligence-Suite
   npm run build
   ```

2. **Excel VBA Setup**
   - Open the built workbook from `dist/` folder
   - Enable Developer tab: File > Options > Customize Ribbon > Developer
   - Set VBA references (Tools > References):
     - Microsoft Scripting Runtime
     - Microsoft Visual Basic for Applications Extensibility 5.3
     - Microsoft XML, v6.0

3. **Optional: OpenAI Integration**
   - Add API key to `%APPDATA%\510k_Tool\openai_key.txt`
   - Get key from https://platform.openai.com/api-keys

## Development Workflow

### Branch Strategy
- `main` - Production-ready code
- `develop` - Integration branch for features
- `feature/*` - Individual feature development
- `hotfix/*` - Critical bug fixes

### Making Changes

1. **Create Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Changes**
   - VBA code: Edit files in `src/vba/`
   - Power Query: Edit `src/powerquery/FDA_510k_Query.pq`
   - Documentation: Update relevant `.md` files

3. **Build and Test**
   ```bash
   npm run build
   # Test the built workbook in dist/ folder
   ```

4. **Commit Changes**
   ```bash
   git add -A
   git commit -m "feat: descriptive commit message"
   ```

5. **Submit Pull Request** (see [Submission Process](#submission-process))

## Code Standards

### VBA Style Guide

```vba
' ==========================================================================
' Module      : mod_YourModule
' Author      : Your Name
' Date        : 2025-01-XX
' Version     : 1.0
' ==========================================================================
' Description : Clear description of module purpose
' Dependencies: List any required modules or references
' ==========================================================================
Option Explicit

Public Function YourFunction(param As String) As Boolean
    ' Brief description of function
    Dim result As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Your code here
    result = True
    
    YourFunction = result
    Exit Function
    
ErrorHandler:
    mod_Logger.LogError "YourFunction", Err.Description
    YourFunction = False
End Function
```

### Code Quality Requirements
-  Use `Option Explicit` in all modules
-  Include comprehensive error handling
-  Add descriptive header comments
-  Use meaningful variable names
-  Follow consistent indentation (4 spaces)
-  Add logging for significant operations
-  Validate all input parameters

### Power Query Standards
```powerquery
// Clear comments explaining data transformations
let
    // Configuration section
    BaseUrl = "https://api.example.com",
    
    // Data retrieval
    Source = Json.Document(Web.Contents(BaseUrl)),
    
    // Data transformation with error handling
    ProcessedData = try Table.FromRecords(Source) otherwise #table({}, {})
in
    ProcessedData
```

## Testing Guidelines

### Manual Testing Checklist
- [ ] Build process completes without errors
- [ ] Excel workbook opens and macros load
- [ ] Power Query refresh works correctly
- [ ] Scoring algorithm produces expected results
- [ ] Error handling works for edge cases
- [ ] Performance is acceptable for typical datasets

### Integration Testing
- Test with real FDA API data
- Verify all VBA modules integrate correctly
- Test Excel object model interactions
- Validate data archiving functionality

### Performance Testing
- Test with large datasets (1000+ records)
- Measure processing time benchmarks
- Verify memory usage stays reasonable
- Test concurrent Power Query operations

## Documentation Standards

### Code Documentation
- Include module headers with purpose and dependencies
- Document all public functions and parameters
- Add inline comments for complex logic
- Update configuration documentation for new settings

### User Documentation
- Update `docs/USER_GUIDE.md` for user-facing changes
- Include screenshots for UI changes
- Document new configuration options
- Update troubleshooting guides

### Technical Documentation
- Update `docs/ARCHITECTURE.md` for structural changes
- Document API changes and integrations
- Include performance impact notes
- Update deployment procedures

## Submission Process

### Before Submitting

1. **Run Quality Checks**
   ```bash
   npm run build        # Verify build works
   npm run clean        # Clean artifacts
   ```

2. **Update Documentation**
   - Update relevant `.md` files
   - Add/update code comments
   - Include screenshots if needed

3. **Test Thoroughly**
   - Manual testing with real data
   - Edge case testing
   - Performance validation

### Pull Request Template

```markdown
## Description
Brief description of changes and motivation.

## Type of Change
- [ ] Bug fix (non-breaking change fixing an issue)
- [ ] New feature (non-breaking change adding functionality)
- [ ] Breaking change (fix or feature causing existing functionality to change)
- [ ] Documentation update

## Testing
- [ ] Manual testing completed
- [ ] Edge cases tested
- [ ] Performance impact assessed
- [ ] Documentation updated

## Screenshots (if applicable)
Include before/after screenshots for UI changes.

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Comments added to complex code
- [ ] Documentation updated
- [ ] No breaking changes (or documented)
```

### Review Process
1. Automated checks must pass
2. Code review by maintainer
3. Testing validation
4. Documentation review
5. Merge approval

## Issue Reporting

### Bug Reports
Include:
- **Environment**: Excel version, Windows version
- **Steps to reproduce**: Detailed step-by-step instructions
- **Expected behavior**: What should happen
- **Actual behavior**: What actually happens
- **Screenshots**: Visual evidence of the issue
- **Error messages**: Exact error text or codes
- **Data context**: Type/size of data being processed

### Feature Requests
Include:
- **Use case**: Why this feature is needed
- **Proposed solution**: How it should work
- **Alternatives considered**: Other approaches
- **Impact**: Who benefits and how
- **Implementation**: Technical considerations

### Priority Labels
- `critical` - System broken, data loss, security issue
- `high` - Major functionality impacted
- `medium` - Minor functionality or enhancement
- `low` - Nice-to-have improvement

## Recognition

Contributors will be recognized in:
- `CHANGELOG.md` for significant contributions
- Code comments for specific implementations
- Project documentation for major features

## Questions?

- Create an issue for technical questions
- Check existing documentation in `docs/` folder
- Review implementation guides for common scenarios

Thank you for contributing to the FDA 510(k) Intelligence Suite! =€