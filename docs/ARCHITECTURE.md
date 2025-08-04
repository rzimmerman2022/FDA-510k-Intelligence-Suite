```
 █████╗ ██████╗  ██████╗██╗  ██╗██╗████████╗███████╗ ██████╗████████╗██╗   ██╗██████╗ ███████╗
██╔══██╗██╔══██╗██╔════╝██║  ██║██║╚══██╔══╝██╔════╝██╔════╝╚══██╔══╝██║   ██║██╔══██╗██╔════╝
███████║██████╔╝██║     ███████║██║   ██║   █████╗  ██║        ██║   ██║   ██║██████╔╝█████╗  
██╔══██║██╔══██╗██║     ██╔══██║██║   ██║   ██╔══╝  ██║        ██║   ██║   ██║██╔══██╗██╔══╝  
██║  ██║██║  ██║╚██████╗██║  ██║██║   ██║   ███████╗╚██████╗   ██║   ╚██████╔╝██║  ██║███████╗
╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝╚═╝  ╚═╝╚═╝   ╚═╝   ╚══════╝ ╚═════╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚══════╝
```

# Architecture Overview - FDA 510(k) Intelligence Suite

> **A comprehensive technical guide for developers and AI assistants working with the FDA 510(k) Intelligence Suite**

## 1. System Overview

The FDA 510(k) Intelligence Suite is an enterprise Excel-based solution that automates the entire workflow of fetching, analyzing, and scoring FDA medical device clearances. Built on a modular VBA architecture with Power Query integration, it provides regulatory intelligence teams with prioritized, actionable insights on recent 510(k) clearances.

### 1.1 Key Value Propositions
- **Automated Data Pipeline**: Zero-touch monthly data refresh from FDA's official API
- **Intelligent Scoring**: Multi-factor algorithm identifies high-value opportunities
- **Company Intelligence**: Cached insights with optional AI-powered summaries
- **Enterprise Ready**: Robust error handling, logging, and maintainer controls

### 1.2 Technical Philosophy
- **Modular Design**: Each VBA module has a single responsibility
- **Error Resilience**: Every procedure includes comprehensive error handling
- **Performance First**: Array processing and minimal worksheet interactions
- **AI Maintainable**: Clear naming, extensive comments, predictable patterns

## 2. Core Technologies

*   **Microsoft Excel:** The primary platform providing the user interface, data storage (sheets, tables), and the VBA execution environment.
*   **Power Query (M Language):** Used for robust data acquisition from the openFDA web API, including dynamic date calculations, JSON parsing, and initial data shaping.
*   **VBA (Visual Basic for Applications):** The core orchestration engine. Handles triggering the workflow, interacting with Power Query, executing the custom scoring algorithm, managing caching, calling external APIs (OpenAI), applying complex formatting, handling archiving, and managing logging.
*   **Excel Tables (ListObjects):** Used on the `Weights` sheet to provide a user-friendly way to configure scoring parameters (weights, keywords) without modifying VBA code.
*   **Scripting Runtime (VBA References):** Used for `Dictionary` objects (efficient lookups) and `FileSystemObject` (API key file reading).
*   **MSXML2.XMLHTTP (VBA References):** Used for making the optional HTTP request to the OpenAI API.
*   **WScript.Shell (VBA References):** Used for expanding environment strings (`%APPDATA%`) in file paths.
*   **VBIDE (VBA References):** Used by `ModuleManager` for interacting with the VBA project itself (exporting/importing components).

## 3. Data Flow Architecture

### 3.1 Visual Flow Diagram
```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  OpenFDA API    │────▶│  Power Query     │────▶│ CurrentMonthData│
│ (Previous Month)│     │  (FDA_510k_Query)│     │    (Sheet)      │
└─────────────────┘     └──────────────────┘     └────────┬────────┘
                                                           │
                        ┌──────────────────────────────────▼────────┐
                        │         VBA Processing Pipeline           │
                        │  ┌────────────┐  ┌──────────────────┐   │
                        │  │   Weights  │  │  Company Cache   │   │
                        │  │  (Tables)  │  │   (Local DB)     │   │
                        │  └──────┬─────┘  └────────┬─────────┘   │
                        │         │                 │              │
                        │  ┌──────▼─────────────────▼──────────┐  │
                        │  │      Scoring Algorithm            │  │
                        │  │  - AC/PC/ST Weights              │  │
                        │  │  - Keyword Matching              │  │
                        │  │  - Negative Factors              │  │
                        │  │  - Synergy Bonuses               │  │
                        │  └────────────┬──────────────────────┘  │
                        └───────────────┼──────────────────────────┘
                                       │
                        ┌──────────────▼──────────────┐
                        │     Formatted Output        │
                        │  - Conditional Formatting   │
                        │  - Smart Column Layout     │
                        │  - Device Name Truncation   │
                        └──────────────┬──────────────┘
                                       │
                        ┌──────────────▼──────────────┐
                        │   Optional: Archive Sheet   │
                        │   (Monthly Static Copy)     │
                        └─────────────────────────────┘
```

### 3.2 Processing Sequence (For AI Understanding)

```vba
' STEP 1: Workbook Opens
Private Sub Workbook_Open()
    Call ProcessMonthly510k  ' Entry point
End Sub

' STEP 2: Main Orchestration
Public Sub ProcessMonthly510k()
    ' 2.1 Initialize environment
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 2.2 Refresh data from FDA
    Call RefreshPowerQuery("Query - pgGet510kData")
    
    ' 2.3 Load configuration
    Call mod_Weights.LoadAll()
    
    ' 2.4 Process each record
    For Each record In dataTable
        score = Calculate510kScore(record)
        recap = GetCompanyRecap(record.Applicant)
        ' Write results back
    Next
    
    ' 2.5 Apply formatting
    Call ApplyConditionalFormatting()
    
    ' 2.6 Archive if needed
    Call ArchiveIfNeeded()
End Sub
```

### 3.3 Key Data Transformations

1. **Raw API Data → Structured Table**
   - JSON response parsed by Power Query
   - Date strings converted to Excel dates
   - Processing time calculated (DecisionDate - ReceivedDate)
   - FDA link constructed from K_Number

2. **Structured Table → Scored Dataset**
   - Each row evaluated against weight tables
   - Keywords matched using array comparison
   - Negative factors applied conditionally
   - Final score calculated and categorized

3. **Scored Dataset → Presentation Layer**
   - Conditional formatting based on score ranges
   - Device names truncated for display
   - Company recaps fetched/generated
   - Archive snapshot created monthly

## 4. Key Components

### 4.1. Excel Workbook (`FDA-510k-Intelligence-Suite.xlsm`)

*   The main container file holding all sheets, code, tables, and connections.

### 4.2. Power Query (`Query - pgGet510kData` or similar)

*   **Source:** `src/powerquery/FDA_510k_Query.pq` (for reference)
*   **Trigger:** Refreshed by VBA (`mod_DataIO.RefreshPowerQuery` function) during the `Workbook_Open` event.
*   **Functionality:**
    *   Calculates the start/end dates for the *previous* full calendar month.
    *   Constructs a dynamic URL for the openFDA 510(k) endpoint, filtering by `decision_date`.
    *   Includes a `limit=1000` parameter (assumes <1000 records/month).
    *   Fetches JSON data via `Web.Contents`.
    *   Handles potential API errors using `try...otherwise`.
    *   Parses the JSON response and expands required fields (`k_number`, `applicant`, `deviceName`, dates, codes, etc.).
    *   Renames fields to user-friendly names.
    *   Performs basic type conversions (Date, Text).
    *   Calculates `ProcTimeDays` (DecisionDate - DateReceived).
    *   Constructs the `FDA_Link` URL.
    *   Selects and reorders a base set of columns.
    *   Sorts data by `DecisionDate` descending.
    *   Loads the resulting table into the `CurrentMonthData` sheet, overwriting previous content.

### 4.3. Excel Sheets

*   **`CurrentMonthData`:**
    *   Target for the Power Query output.
    *   Contains the main data table (`ListObject`).
    *   VBA adds calculated columns directly to this table (scores, weights, recap, category).
    *   Final formatting is applied here by `mod_Format`.
*   **`Weights`:**
    *   Contains user-configurable Excel Tables (`ListObjects`):
        *   `tblACWeights` (AC Code -> Weight)
        *   `tblSTWeights` (Submission Type -> Weight)
        *   `tblPCWeights` (Product Code -> Weight - Optional)
        *   `tblKeywords` (High-value keywords for scoring)
        *   `tblNFCosmeticKeywords` (Keywords triggering cosmetic negative factor)
        *   `tblNFDiagnosticKeywords` (Keywords triggering diagnostic negative factor)
        *   `tblTherapeuticKeywords` (Keywords potentially negating negative factors)
    *   *(Optional)* May contain the `DebugMode` Named Range cell used by `mod_Logger`.
*   **`CompanyCache`:**
    *   Acts as a simple, persistent database for company summaries.
    *   Columns: `CompanyName`, `RecapText`, `LastUpdated`.
    *   Managed by `mod_Cache`.
*   **`RunLog` (Hidden - `xlSheetVeryHidden`):**
    *   Created/managed by the `mod_Logger` module.
    *   Stores a persistent history of processing runs, events, warnings, and errors.
    *   Columns: `RunID`, `Timestamp`, `User`, `Step`, `Level`, `Message`, `Extra`.
*   **`DebugTrace` (Hidden - `xlSheetHidden` or `xlSheetVeryHidden`):**
    *   Created/managed by the `mod_DebugTraceHelpers` module.
    *   Stores verbose trace messages if enabled via constants in that module.
    *   Columns: `Timestamp`, `Level`, `Procedure`, `Message`, `Details`.

### 4.4. VBA Modules (Gold Standard Organization)

**IMPORTANT FOR AI ASSISTANTS**: Each module follows strict patterns:
- `Option Explicit` always at top
- Error handlers in every public procedure
- Cleanup code in dedicated labels
- Module-level constants for configuration
- Clear separation of public API vs private implementation

#### **Core Business Logic** (`src/vba/core/`)
*   **`mod_510k_Processor.bas`:**
    *   Main processing orchestration and workflow management
    *   `ProcessMonthly510k()`: Primary entry point for data processing pipeline
*   **`mod_Archive.bas`:**
    *   Monthly archive sheet creation and management
    *   `ArchiveIfNeeded()`: Copies data sheet, converts to static values
*   **`mod_Cache.bas`:**
    *   Company recap caching system with OpenAI integration
    *   `LoadCompanyCache()`, `SaveCompanyCache()`, `GetCompanyRecap()`
*   **`mod_Schema.bas`:**
    *   Table structure management and safe data access
    *   `GetColumnIndices()`, `SafeGetString()`, `SafeGetVariant()`
*   **`mod_Score.bas`:**
    *   Core FDA 510(k) scoring algorithm implementation
    *   `Calculate510kScore()`: Multi-factor scoring with configurable weights
*   **`mod_Weights.bas`:**
    *   Scoring parameter loading and management from Excel tables
    *   `LoadAll()`, `GetACWeights()`, `GetSTWeights()`, `GetPCWeights()`

#### **Shared Utilities** (`src/vba/utilities/`)
*   **`mod_Config.bas`:**
    *   Global configuration constants and system settings
    *   Maintainer settings, API paths, sheet names, defaults
*   **`mod_DataIO.bas`:**
    *   Data input/output operations and Power Query management
    *   `RefreshPowerQuery()`, `SheetExists()`, `ArrayToTable()`
*   **`mod_Format.bas`:**
    *   Visual formatting, styling, and UI operations
    *   `ApplyAll()`, column management, conditional formatting
*   **`mod_Utils.bas`:**
    *   Miscellaneous helper functions and utilities
    *   `GetWorksheets()`, `IsMaintainerUser()`, `EnsureUIOn()`
*   **Debug & Logging Modules:**
    *   `mod_Debug.bas`: Legacy debugging utilities
    *   `mod_Logger.bas`: Buffered logging to RunLog sheet
    *   `mod_DebugColumnTrace.bas`: Column-specific debugging
    *   `mod_DebugTraceHelpers.bas`: Verbose trace system
    *   `mod_DirectTrace.bas`: Direct trace utilities
    *   `mod_ColumnDebugger.bas`: Column debugging tools
    *   `StandaloneDebug.bas`: Standalone debug functions

#### **Application Modules** (`src/vba/modules/`)
*   **`ThisWorkbook.cls`:**
    *   Workbook event handlers (Open, BeforeClose)
    *   Primary application entry points
*   **`mod_RefreshSolutions.bas`:**
    *   Power Query refresh solutions and connection management
*   **`mod_TestRefresh.bas`:**
    *   Test utilities for refresh operations
*   **`mod_TestWithContext.bas`:**
    *   Context-aware testing functionality
*   **`ModuleManager.bas`:**
    *   VBA code export/import for version control

## 4. Data Flow & Workflow (Updated)

1.  **Workbook Open (`ThisWorkbook.Workbook_Open`):**
    *   Application state optimized (ScreenUpdating Off, etc.).
    *   `mod_DataIO.RefreshPowerQuery()` is called -> Power Query fetches data for the previous month -> Data loaded into `CurrentMonthData` table.
    *   `mod_510k_Processor.ProcessMonthly510k()` is called.
2.  **Main Processing (`mod_510k_Processor.ProcessMonthly510k`):**
    *   **Guard Check:** Determines if full processing should run (missing archive OR days 1-5 OR maintainer (`mod_Utils.IsMaintainerUser`)). If not, jumps to Archive Check.
    *   **Setup:** Get sheet objects (`mod_Utils.GetWorksheets`), ensure columns exist (`mod_Format.AddScoreColumnsIfNeeded`), map column indices (`mod_Schema.GetColumnIndices`).
    *   **Load Parameters:** `mod_Weights.LoadAll()` reads `Weights` sheet tables. `mod_Cache.LoadCompanyCache()` reads `CompanyCache` sheet.
    *   **Read Data:** `tblData.DataBodyRange.Value2` copied into `dataArr` variant array.
    *   **Processing Loop:** Iterate through `dataArr`:
        *   Call `mod_Score.Calculate510kScore()` using current row data, column map (`mod_Schema`), loaded weights/keywords (`mod_Weights`), and rules/defaults (`mod_Config`).
        *   Call `mod_Cache.GetCompanyRecap()` (checks memory, optionally calls OpenAI if maintainer & needed).
        *   Update the `dataArr` in memory with scores/recap.
    *   **Write Back:** Call `mod_DataIO.ArrayToTable()` to write modified `dataArr` back to the `CurrentMonthData` table.
    *   **Save Cache:** Call `mod_Cache.SaveCompanyCache()` to write memory cache back to the `CompanyCache` sheet.
    *   **Format:** Call `mod_Format.ApplyAll()` which orchestrates deleting duplicates, applying number formats, table style, colors, reordering columns, sorting, and autofitting.
    *   **(Jump Target `CleanExit` includes Archive Check):** Calculate expected `archiveSheetName`. If `mod_DataIO.SheetExists()` returns false, call `mod_Archive.ArchiveIfNeeded()`.
    *   **Logging:** `mod_Logger.LogEvt` and `mod_DebugTraceHelpers.TraceEvt` called throughout.
    *   **Cleanup:** Call `mod_Utils.EnsureUIOn()` to restore application settings, release local object variables.
3.  **Workbook Close (`ThisWorkbook.Workbook_BeforeClose`):**
    *   Call `mod_Logger.TrimRunLog()`.

## 5. Configuration Points (Updated)

*   **VBA Constants in `mod_Config`:** This is the **primary** location for configuration (Maintainer name, file paths, sheet/table names, default scores, API settings, version info).
*   **Excel Tables on `Weights` sheet:** User-editable weights and keywords.
*   **Named Range `DebugMode`:** Optional control for `mod_Logger` detail level (read by `mod_Logger.DebugModeOn`).
*   **Constants in `mod_DebugTraceHelpers`:** `TRACE_ENABLED` and `TRACE_LEVEL` control the verbose tracing system.
*   **API Key File:** External text file containing the OpenAI key (path defined in `mod_Config`).

## 6. Dependencies

*   Excel with VBA & Power Query.
*   Windows OS (assumed).
*   Internet connection.
*   VBA References: `Microsoft Scripting Runtime`, `Microsoft XML, v6.0`, `Microsoft Visual Basic for Applications Extensibility 5.3`.
*   *(Optional)* OpenAI Account & API Key.

## 7. Error Handling

*   Uses standard VBA `On Error GoTo <Label>` within major routines.
*   Specific error handlers log details using `mod_Logger.LogEvt` and/or `mod_DebugTraceHelpers.TraceEvt`.
*   User-facing `MsgBox` used for critical errors preventing execution or major warnings.
*   `try...otherwise` used in Power Query for API call robustness.
*   `On Error Resume Next` used sparingly for non-critical checks or operations where failure is acceptable (e.g., optional table loading, UI operations like autofit).

## 8. Limitations & Considerations

*   **API Rate Limits:** Assumes usage stays within openFDA and OpenAI API limits. No specific rate-limiting code implemented.
*   **PQ Record Limit:** Assumes < 1000 records per month from FDA API. Pagination not implemented.
*   **OpenAI Cost/Reliability:** OpenAI calls incur costs and depend on API availability. The basic JSON parsing is fragile.
*   **Scalability:** Performance might degrade with extremely large numbers of keywords or cache entries, although array processing helps significantly.
*   **Cross-Platform:** Developed/tested primarily on Windows; Mac compatibility may vary (especially regarding `Environ`, `WScript.Shell`, `MSXML`, file paths, VBE Extensibility).
*   **Hardcoded Paths in `ModuleManager`:** The `ModuleManager.bas` utility uses hardcoded paths that need manual adjustment if used.

## 9. AI Assistant Development Guide

### 9.1 Common Code Patterns

#### Error Handling Pattern
```vba
Public Sub StandardProcedure()
    On Error GoTo ErrorHandler
    Dim cleanup As Boolean: cleanup = False
    
    ' Main logic here
    cleanup = True
    
CleanExit:
    If cleanup Then
        ' Cleanup code
        Set obj = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    LogEvt "Error in StandardProcedure", lgERROR, Err.Description
    Resume CleanExit
End Sub
```

#### Array Processing Pattern
```vba
' GOOD: Process in memory
Dim dataArr As Variant
dataArr = ws.Range("A1:Z1000").Value

For i = 1 To UBound(dataArr, 1)
    ' Process dataArr(i, columnIndex)
Next i

ws.Range("A1:Z1000").Value = dataArr

' BAD: Cell-by-cell access
For i = 1 To 1000
    ws.Cells(i, 1).Value = ProcessValue(ws.Cells(i, 2).Value)
Next i
```

#### Dictionary Usage Pattern
```vba
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
dict.CompareMode = vbTextCompare ' Case-insensitive

' Add items
If Not dict.Exists(key) Then
    dict.Add key, value
End If

' Clean up
Set dict = Nothing
```

### 9.2 Module Interaction Rules

1. **Never bypass the public API** - Always call public procedures
2. **Respect module boundaries** - Don't access private variables
3. **Use proper cleanup** - Set objects to Nothing
4. **Log all errors** - Use LogEvt for error tracking
5. **Test with arrays** - Minimize worksheet interactions

### 9.3 Performance Guidelines

| Operation | Good Practice | Bad Practice |
|-----------|--------------|--------------|
| Reading Data | `arr = Range.Value` | `For Each Cell In Range` |
| Writing Data | `Range.Value = arr` | `Cell.Value = x` (in loop) |
| Finding Data | `Application.Match()` | Loop through cells |
| Sorting | Use Excel's built-in sort | Bubble sort in VBA |

### 9.4 Debugging Tips

1. **Enable debug mode**: Set `DEBUG_MODE = True` in mod_Config
2. **Check RunLog sheet**: All errors are logged there
3. **Use immediate window**: `Debug.Print` key variables
4. **Step through code**: F8 in VBA editor
5. **Check array bounds**: Common source of errors

### 9.5 Common Pitfalls to Avoid

- **Don't assume sheet exists** - Always check with error handling
- **Don't hardcode paths** - Use configuration constants
- **Don't skip cleanup** - Memory leaks crash Excel
- **Don't ignore errors** - Log and handle gracefully
- **Don't mix data types** - Variant arrays need careful handling
