# Configuration Files

This directory contains configuration files for the FDA 510(k) Intelligence Suite. These files are for reference and documentation purposes - the actual configuration is embedded within the Excel VBA code.

## Files

### app.config.json
Main application configuration including:
- API endpoints and settings
- Scoring weights and thresholds
- Excel sheet and table names
- Logging configuration
- Feature flags
- Performance settings

### environment.json
Environment-specific settings for:
- Development
- Staging  
- Production

## Usage

These JSON files document the configuration structure used by the application. To modify settings:

1. Open the Excel workbook
2. Press Alt+F11 to open VBA editor
3. Navigate to `mod_Config` module
4. Update the relevant constants

## Note

The VBA application does not read these JSON files directly. They serve as:
- Documentation of available settings
- Reference for developers
- Template for future JSON-based configuration