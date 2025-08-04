# FDA 510(k) Intelligence Suite - Pipeline Summary

## Overview
The FDA 510(k) Intelligence Suite is a sophisticated Excel-based data pipeline that automatically fetches, analyzes, and scores FDA medical device clearances to identify high-value opportunities for regulatory intelligence teams.

## Main Pipeline Flow

### 1. Data Acquisition (Power Query)
- **Source**: OpenFDA API (https://api.fda.gov/device/510k.json)
- **Frequency**: Automatic on workbook open
- **Time Range**: Previous calendar month
- **Query**: `src/powerquery/FDA_510k_Query.pq`
- **Output**: ~200-500 records per month to CurrentMonthData sheet

### 2. VBA Processing Engine
The core orchestrator (`mod_510k_Processor.bas`) manages the entire workflow:

#### a. Data Refresh
- Power Query refresh via `mod_DataIO`
- Connection management and error handling
- Automatic retry logic for API failures

#### b. Parameter Loading
- Advisory Committee (AC) weights from `tblACWeights`
- Product Code (PC) weights from `tblPCWeights`
- Submission Type (ST) weights from `tblSTWeights`
- Keyword lists for scoring and negative factors

#### c. Scoring Algorithm (`mod_Score`)
Multi-factor scoring with configurable weights:
- **Base Score**: AC weight + PC weight + ST weight
- **Keyword Bonus**: +2-5 points per matching keyword
- **Processing Time**: Faster approvals get bonus points
- **Geographic Factor**: US-based companies preferred
- **Negative Factors**: -5 to -10 for cosmetic/diagnostic keywords
- **Synergy Bonuses**: +10 for specific keyword combinations

#### d. Company Intelligence (`mod_Cache`)
- Local caching system for company summaries
- Optional OpenAI integration for AI-generated recaps
- Persistent storage in CompanyCache sheet

#### e. Formatting (`mod_Format`)
- Conditional formatting based on score ranges
- Smart column management and device name truncation
- Professional UI with color-coded priorities

#### f. Archiving (`mod_Archive`)
- Monthly snapshots with static values
- Automatic archive creation on new month
- Historical data preservation

### 3. Output
- **Primary**: CurrentMonthData sheet with scored, formatted results
- **Secondary**: Monthly archive sheets (e.g., "2025-07 Archive")
- **Logging**: RunLog sheet tracks all processing events

## Key Features

### Automated Operation
- Triggers on workbook open
- No user intervention required
- Robust error handling throughout

### Enterprise Ready
- Comprehensive logging system
- Maintainer bypass controls
- Performance optimizations
- Modular architecture for easy maintenance

### Configurable
- All weights manageable via Excel tables
- No code changes needed for scoring adjustments
- Debug modes for troubleshooting

## Technical Stack
- **Excel 2016+**: Primary platform
- **Power Query (M)**: Data acquisition
- **VBA**: Processing logic
- **OpenFDA API**: Data source
- **OpenAI API**: Optional AI summaries

## Repository Structure (Cleaned)
```
FDA-510k-Intelligence-Suite/
├── assets/excel-workbooks/     # Main Excel files
├── config/                     # Configuration files
├── docs/                       # Documentation
├── samples/                    # Sample data
├── scripts/                    # Build/deploy scripts
├── src/                        # Source code
│   ├── powerquery/            # Power Query files
│   └── vba/                   # VBA modules
│       ├── core/              # Business logic
│       ├── modules/           # App modules
│       └── utilities/         # Helpers
└── tests/                      # Test files
```

## Build & Deployment
```bash
# Build for distribution
npm run build

# Deploy to target
scripts\deploy\deploy.bat "C:\Target\Path"
```

## Maintenance Notes
- Legacy VBA code removed from `src/vba/_legacy`
- Working documents moved to `.archives`
- Repository structure follows enterprise standards
- All documentation updated to reflect current pipeline