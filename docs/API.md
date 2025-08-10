# VBA API Reference - FDA 510(k) Intelligence Suite

**Version**: 1.9.0  
**Last Updated**: 2025-08-10  
**Description**: Complete API reference for all public VBA procedures and functions

## Overview

This document provides a comprehensive API reference for all public VBA procedures and functions in the FDA 510(k) Intelligence Suite. The API is organized by module and includes parameters, return values, and usage examples.

## Table of Contents

1. [Core Modules](#core-modules)
   - [mod_510k_Processor](#mod_510k_processor)
   - [mod_Score](#mod_score)
   - [mod_Cache](#mod_cache)
   - [mod_Archive](#mod_archive)
   - [mod_Weights](#mod_weights)
   - [mod_Schema](#mod_schema)

2. [Application Modules](#application-modules)
   - [ThisWorkbook](#thisworkbook)
   - [ModuleManager](#modulemanager)

3. [Utility Modules](#utility-modules)
   - [mod_DataIO](#mod_dataio)
   - [mod_Logger](#mod_logger)
   - [mod_Format](#mod_format)
   - [mod_Utils](#mod_utils)
   - [mod_Config](#mod_config)

---

## Core Modules

### mod_510k_Processor

**Purpose**: Central orchestrator for the FDA 510(k) processing pipeline

#### Public Procedures

##### `ProcessMonthly510k()`
**Description**: Main entry point that executes the complete 510(k) processing pipeline

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_510k_Processor.ProcessMonthly510k()
```

**Process Flow**:
1. Initialize logging and error handling
2. Refresh Power Query data connection
3. Load scoring weights and parameters
4. Calculate scores for each record
5. Apply formatting and company intelligence
6. Archive data if new month detected

**Dependencies**:
- mod_DataIO.RefreshPowerQuery
- mod_Weights.LoadAllWeights
- mod_Score.Calculate510kScore
- mod_Cache.ProcessCompanyCache
- mod_Format.ApplyFormatting
- mod_Archive.ArchiveIfNewMonth

---

### mod_Score

**Purpose**: Implements the multi-factor scoring algorithm

#### Public Functions

##### `Calculate510kScore(record As Variant) As Long`
**Description**: Calculates the overall score for a single 510(k) record

**Parameters**:
- `record` (Variant): Array containing the 510(k) record data

**Returns**: Long - Calculated score value

**Usage**:
```vba
Dim score As Long
score = mod_Score.Calculate510kScore(recordArray)
```

**Scoring Factors**:
- Advisory Committee (AC) weight
- Product Code (PC) weight
- Submission Type (ST) weight
- Keyword matching bonus
- Processing time bonus
- Geographic factors
- Negative keyword penalties

---

### mod_Cache

**Purpose**: Manages company intelligence caching system

#### Public Procedures

##### `ProcessCompanyCache()`
**Description**: Updates company cache with new records and AI summaries

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Cache.ProcessCompanyCache()
```

##### `ClearCache()`
**Description**: Clears all cached company data

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Cache.ClearCache()
```

#### Public Functions

##### `GetCompanySummary(companyName As String) As String`
**Description**: Retrieves cached company summary or generates new one

**Parameters**:
- `companyName` (String): Name of the company to look up

**Returns**: String - Company summary text

**Usage**:
```vba
Dim summary As String
summary = mod_Cache.GetCompanySummary("Acme Medical Devices")
```

---

### mod_Archive

**Purpose**: Handles monthly data archiving functionality

#### Public Procedures

##### `ArchiveIfNewMonth()`
**Description**: Checks if new month detected and creates archive if needed

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Archive.ArchiveIfNewMonth()
```

##### `CreateArchive(archiveDate As Date)`
**Description**: Creates a monthly archive sheet for specified date

**Parameters**:
- `archiveDate` (Date): Date for the archive (typically previous month)

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Archive.CreateArchive(DateSerial(2025, 7, 1))
```

---

### mod_Weights

**Purpose**: Manages scoring weights and keyword lists

#### Public Procedures

##### `LoadAllWeights()`
**Description**: Loads all scoring weights from Excel tables

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Weights.LoadAllWeights()
```

#### Public Functions

##### `GetACWeight(advisoryCommittee As String) As Long`
**Description**: Gets weight value for Advisory Committee

**Parameters**:
- `advisoryCommittee` (String): Advisory committee code

**Returns**: Long - Weight value

**Usage**:
```vba
Dim weight As Long
weight = mod_Weights.GetACWeight("CV")  ' Cardiovascular
```

##### `GetPCWeight(productCode As String) As Long`
**Description**: Gets weight value for Product Code

**Parameters**:
- `productCode` (String): FDA product code

**Returns**: Long - Weight value

**Usage**:
```vba
Dim weight As Long
weight = mod_Weights.GetPCWeight("DQA")
```

##### `GetSTWeight(submissionType As String) As Long`
**Description**: Gets weight value for Submission Type

**Parameters**:
- `submissionType` (String): Type of FDA submission

**Returns**: Long - Weight value

**Usage**:
```vba
Dim weight As Long
weight = mod_Weights.GetSTWeight("510(k)")
```

---

### mod_Schema

**Purpose**: Defines data structures and column mappings

#### Public Constants

##### Column Index Constants
```vba
Public Const COL_K_NUMBER As Long = 1
Public Const COL_DECISION_DATE As Long = 2
Public Const COL_DATE_RECEIVED As Long = 3
Public Const COL_PROC_TIME_DAYS As Long = 4
Public Const COL_APPLICANT As Long = 5
Public Const COL_CONTACT As Long = 6
Public Const COL_DEVICE_NAME As Long = 7
Public Const COL_STATEMENT As Long = 8
Public Const COL_AC As Long = 9
Public Const COL_PC As Long = 10
Public Const COL_SUBM_TYPE As Long = 11
Public Const COL_CITY As Long = 12
Public Const COL_STATE As Long = 13
Public Const COL_COUNTRY As Long = 14
Public Const COL_FDA_LINK As Long = 15
Public Const COL_510K_SCORE As Long = 16
Public Const COL_AC_WEIGHT As Long = 17
Public Const COL_PC_WEIGHT As Long = 18
Public Const COL_ST_WEIGHT As Long = 19
Public Const COL_COMPANY_RECAP As Long = 20
Public Const COL_SCORE_CATEGORY As Long = 21
```

---

## Application Modules

### ThisWorkbook

**Purpose**: Workbook-level event handlers

#### Event Handlers

##### `Workbook_Open()`
**Description**: Executes when workbook opens (auto-processing currently commented out)

**Parameters**: None

**Returns**: Nothing (Event handler)

**Note**: Auto-processing is currently disabled. Uncomment code to enable.

##### `Workbook_BeforeClose(Cancel As Boolean)`
**Description**: Executes before workbook closes, performs cleanup

**Parameters**:
- `Cancel` (Boolean): Whether to cancel the close operation

**Returns**: Nothing (Event handler)

---

### ModuleManager

**Purpose**: VBA code management utilities

#### Public Procedures

##### `ExportAllModules()`
**Description**: Exports all VBA modules to source files

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call ModuleManager.ExportAllModules()
```

##### `ImportAllModules()`
**Description**: Imports VBA modules from source files

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call ModuleManager.ImportAllModules()
```

---

## Utility Modules

### mod_DataIO

**Purpose**: Data input/output operations

#### Public Procedures

##### `RefreshPowerQuery()`
**Description**: Refreshes the Power Query connection to fetch latest data

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_DataIO.RefreshPowerQuery()
```

##### `WriteDataToSheet(data As Variant, sheetName As String)`
**Description**: Writes array data to specified worksheet

**Parameters**:
- `data` (Variant): 2D array of data to write
- `sheetName` (String): Target worksheet name

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_DataIO.WriteDataToSheet(dataArray, "CurrentMonthData")
```

---

### mod_Logger

**Purpose**: Comprehensive logging system

#### Public Procedures

##### `LogEvt(message As String, level As LogLevel, Optional details As String)`
**Description**: Logs an event with specified level and optional details

**Parameters**:
- `message` (String): Main log message
- `level` (LogLevel): Logging level (lgINFO, lgWARN, lgERROR, lgDEBUG)
- `details` (String, Optional): Additional details

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Logger.LogEvt("Processing started", lgINFO)
Call mod_Logger.LogEvt("Connection failed", lgERROR, "Timeout after 30 seconds")
```

##### `FlushLogBuf()`
**Description**: Flushes log buffer to worksheet

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Logger.FlushLogBuf()
```

##### `TrimRunLog()`
**Description**: Trims old entries from RunLog sheet

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Logger.TrimRunLog()
```

---

### mod_Format

**Purpose**: UI formatting and styling

#### Public Procedures

##### `ApplyFormatting()`
**Description**: Applies all formatting rules to the current data

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Format.ApplyFormatting()
```

##### `FormatScoreColumn()`
**Description**: Applies conditional formatting to score column

**Parameters**: None

**Returns**: Nothing (Sub procedure)

**Usage**:
```vba
Call mod_Format.FormatScoreColumn()
```

---

### mod_Utils

**Purpose**: General utility functions

#### Public Functions

##### `GetWorksheetSafe(sheetName As String) As Worksheet`
**Description**: Safely gets worksheet reference with error handling

**Parameters**:
- `sheetName` (String): Name of the worksheet

**Returns**: Worksheet - Worksheet object or Nothing if not found

**Usage**:
```vba
Dim ws As Worksheet
Set ws = mod_Utils.GetWorksheetSafe("CurrentMonthData")
If Not ws Is Nothing Then
    ' Worksheet exists, proceed with operations
End If
```

##### `SafeConvertToLong(value As Variant) As Long`
**Description**: Safely converts variant to Long with error handling

**Parameters**:
- `value` (Variant): Value to convert

**Returns**: Long - Converted value or 0 if conversion fails

**Usage**:
```vba
Dim numValue As Long
numValue = mod_Utils.SafeConvertToLong(cellValue)
```

---

### mod_Config

**Purpose**: Global constants and configuration

#### Public Constants

##### Application Constants
```vba
Public Const APP_NAME As String = "FDA 510(k) Intelligence Suite"
Public Const VERSION_INFO As String = "v1.9.0"
Public Const AUTHOR As String = "Ryan Zimmerman"
```

##### Sheet Names
```vba
Public Const SHEET_CURRENT_MONTH As String = "CurrentMonthData"
Public Const SHEET_COMPANY_CACHE As String = "CompanyCache"
Public Const SHEET_WEIGHTS As String = "Weights"
Public Const SHEET_RUN_LOG As String = "RunLog"
```

##### Table Names
```vba
Public Const TBL_AC_WEIGHTS As String = "tblACWeights"
Public Const TBL_PC_WEIGHTS As String = "tblPCWeights"
Public Const TBL_ST_WEIGHTS As String = "tblSTWeights"
Public Const TBL_KEYWORDS As String = "tblKeywords"
```

#### Public Enums

##### LogLevel
```vba
Public Enum LogLevel
    lgDEBUG = 1
    lgINFO = 2
    lgWARN = 3
    lgERROR = 4
End Enum
```

---

## Error Handling Patterns

All public procedures follow this error handling pattern:

```vba
Public Sub ExampleProcedure()
    On Error GoTo ErrorHandler
    
    ' Main logic here
    
CleanExit:
    ' Cleanup code here
    Exit Sub
    
ErrorHandler:
    mod_Logger.LogEvt "Error in ExampleProcedure", lgERROR, Err.Description
    Resume CleanExit
End Sub
```

## Performance Considerations

1. **Array Processing**: All data operations use arrays to minimize worksheet interactions
2. **Batch Updates**: Multiple changes are applied in single operations
3. **Memory Management**: Objects are explicitly set to Nothing in cleanup sections
4. **Connection Management**: Power Query connections are properly closed after use

## Testing

### Unit Tests
- Located in `tests/unit/`
- Test individual functions with known inputs/outputs
- Run via Excel VBA Test Suite

### Integration Tests
- Located in `tests/integration/`
- Test complete workflows end-to-end
- Verify data pipeline functionality

---

**For additional information, see:**
- [Architecture Guide](ARCHITECTURE.md) - System design details
- [Development Guide](AI_DEVELOPMENT_GUIDE.md) - Coding standards
- [User Guide](USER_GUIDE.md) - End-user instructions