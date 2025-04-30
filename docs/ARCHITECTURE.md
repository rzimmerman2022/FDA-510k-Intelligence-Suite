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

### 3.4. VBA Modules (Refactored Structure)

*   **`ThisWorkbook` (Class Module):**
    *   Source: `src/vba/ThisWorkbook.cls`
    *   Handles workbook events.
    *   `Workbook_Open()`: Entry point, calls `mod_DataIO.RefreshPowerQuery` then `mod_510k_Processor.ProcessMonthly510k`.
    *   `Workbook_BeforeClose()`: Calls `mod_Logger.TrimRunLog`.
*   **`mod_Config`:**
    *   Source: `src/vba/mod_Config.bas`
    *   Central repository for **all** global `Public Const` values (sheet names, API paths, scoring defaults, version info, maintainer username, etc.). No procedural code.
*   **`mod_DataIO`:**
    *   Source: `src/vba/mod_DataIO.bas`
    *   Handles data input/output.
    *   `RefreshPowerQuery()`: Manages the Power Query refresh process, including enabling/disabling refresh.
    *   `SheetExists()`: Checks if a sheet exists.
    *   `CleanupDuplicateConnections()`: Removes duplicate PQ connections after sheet copy.
    *   `ArrayToTable()`: Writes VBA array data back to an Excel table.
*   **`mod_Schema`:**
    *   Source: `src/vba/mod_Schema.bas`
    *   Manages table structure understanding.
    *   `GetColumnIndices()`: Creates map of header names to column indices, handles duplicates, checks required columns.
    *   `SafeGetString()`, `SafeGetVariant()`, `SafeGetColIndex()`: Safely access data/indices using the column map.
    *   `ColumnExistsInMap()`: Helper to check for column existence in the map.
*   **`mod_Weights`:**
    *   Source: `src/vba/mod_Weights.bas`
    *   Loads and provides access to scoring parameters.
    *   `LoadAll()`: Reads all tables from the `Weights` sheet into memory (dictionaries/collections).
    *   `GetACWeights()`, `GetSTWeights()`, `GetPCWeights()`, `GetHighValueKeywords()`, etc.: Public functions to access the loaded data.
*   **`mod_Score`:**
    *   Source: `src/vba/mod_Score.bas`
    *   Contains the core scoring algorithm.
    *   `Calculate510kScore()`: Calculates score for one row using data from `mod_Schema`, weights/keywords from `mod_Weights`, and defaults/rules from `mod_Config`. Uses `CheckKeywords` helper.
*   **`mod_Cache`:**
    *   Source: `src/vba/mod_Cache.bas`
    *   Manages the company recap cache.
    *   `LoadCompanyCache()`, `SaveCompanyCache()`: Interact with the `CompanyCache` sheet.
    *   `GetCompanyRecap()`: Retrieves recap from memory or calls `GetCompanyRecapOpenAI()`.
    *   `GetCompanyRecapOpenAI()`: Handles optional OpenAI API call (requires API key file defined in `mod_Config`).
*   **`mod_Format`:**
    *   Source: `src/vba/mod_Format.bas`
    *   Applies all visual formatting to the `CurrentMonthData` table.
    *   `ApplyAll()`: Orchestrates calls to private helpers.
    *   Helpers: `DeleteDuplicateColumns()`, `ApplyNumberFormats()`, `FormatTableLook()`, `FormatCategoryColors()`, `ReorganizeColumns()`, `SortDataTable()`, `FreezeHeaderAndFirstColumns()` (currently disabled).
*   **`mod_Archive`:**
    *   Source: `src/vba/mod_Archive.bas`
    *   Handles creation of monthly archive sheets.
    *   `ArchiveIfNeeded()`: Copies the data sheet, renames it, converts table to range, calls `mod_DataIO.CleanupDuplicateConnections`.
*   **`mod_Logger`:**
    *   Source: `src/vba/mod_Logger.bas`
    *   Provides buffered logging to the `RunLog` sheet.
    *   `LogEvt()`, `FlushLogBuf()`, `TrimRunLog()`. Includes `DebugModeOn()` check based on maintainer status (`mod_Utils`) and `DebugMode` named range.
*   **`mod_DebugTraceHelpers`:**
    *   Source: `src/vba/mod_DebugTraceHelpers.bas`
    *   Provides conditional, verbose tracing to the `DebugTrace` sheet. Controlled by constants within the module.
    *   `TraceEvt()`, `ClearDebugTrace()`.
*   **`mod_Debug`:**
    *   Source: `src/vba/mod_Debug.bas`
    *   Contains older/simpler debugging utilities (`DumpHeaders`, `DebugTrace`). Potential candidate for refactoring/removal.
*   **`mod_Utils`:**
    *   Source: `src/vba/mod_Utils.bas`
    *   Contains miscellaneous helper functions.
    *   `GetWorksheets()`, `IsMaintainerUser()`, `EnsureUIOn()`, `GetBrightness()`.
*   **`ModuleManager`:**
    *   Source: `src/vba/ModuleManager.bas`
    *   Provides utilities for exporting/importing VBA code modules for version control. Requires VBE references.

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
