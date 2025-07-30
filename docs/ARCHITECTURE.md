# Architecture Overview - FDA 510(k) Intelligence Suite

## 1. High-Level Goal

This tool automates the retrieval, scoring, and presentation of FDA 510(k) clearance data within a single Microsoft Excel workbook (`.xlsm`). The primary goal is to provide actionable intelligence on recent clearances based on a configurable scoring model, minimizing manual data handling and analysis.

## 2. Core Technologies

*   **Microsoft Excel:** The primary platform providing the user interface, data storage (sheets, tables), and the VBA execution environment.
*   **Power Query (M Language):** Used for robust data acquisition from the openFDA web API, including dynamic date calculations, JSON parsing, and initial data shaping.
*   **VBA (Visual Basic for Applications):** The core orchestration engine. Handles triggering the workflow, interacting with Power Query, executing the custom scoring algorithm, managing caching, calling external APIs (OpenAI), applying complex formatting, handling archiving, and managing logging.
*   **Excel Tables (ListObjects):** Used on the `Weights` sheet to provide a user-friendly way to configure scoring parameters (weights, keywords) without modifying VBA code.
*   **Scripting Runtime (VBA References):** Used for `Dictionary` objects (efficient lookups) and `FileSystemObject` (API key file reading).
*   **MSXML2.XMLHTTP (VBA References):** Used for making the optional HTTP request to the OpenAI API.
*   **WScript.Shell (VBA References):** Used for expanding environment strings (`%APPDATA%`) in file paths.
*   **VBIDE (VBA References):** Used by `ModuleManager` for interacting with the VBA project itself (exporting/importing components).

## 3. Key Components

### 3.1. Excel Workbook (`FDA-510k-Intelligence-Suite.xlsm`)

*   The main container file holding all sheets, code, tables, and connections.

### 3.2. Power Query (`Query - pgGet510kData` or similar)

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

### 3.3. Excel Sheets

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

### 3.4. VBA Modules (Gold Standard Organization)

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
