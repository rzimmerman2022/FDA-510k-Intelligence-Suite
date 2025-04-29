# Architecture Overview - FDA 510(k) Intelligence Suite

## 1. High-Level Goal

This tool automates the retrieval, scoring, and presentation of FDA 510(k) clearance data within a single Microsoft Excel workbook (`.xlsm`). The primary goal is to provide actionable intelligence on recent clearances based on a configurable scoring model, minimizing manual data handling and analysis.

## 2. Core Technologies

* **Microsoft Excel:** The primary platform providing the user interface, data storage (sheets, tables), and the VBA execution environment.
* **Power Query (M Language):** Used for robust data acquisition from the openFDA web API, including dynamic date calculations, JSON parsing, and initial data shaping.
* **VBA (Visual Basic for Applications):** The core orchestration engine. Handles triggering the workflow, interacting with Power Query, executing the custom scoring algorithm, managing caching, calling external APIs (OpenAI), applying complex formatting, handling archiving, and managing logging.
* **Excel Tables (ListObjects):** Used on the `Weights` sheet to provide a user-friendly way to configure scoring parameters (weights, keywords) without modifying VBA code.
* **Scripting Runtime (VBA References):** Used for `Dictionary` objects (efficient lookups) and `FileSystemObject` (API key file reading).
* **MSXML2.XMLHTTP (VBA References):** Used for making the optional HTTP request to the OpenAI API.
* **WScript.Shell (VBA References):** Used for expanding environment strings (`%APPDATA%`) in file paths.

## 3. Key Components

### 3.1. Excel Workbook (`FDA-510k-Intelligence-Suite.xlsm`)

* The main container file holding all sheets, code, tables, and connections.

### 3.2. Power Query (`Query - pqGet510kData`)

* **Source:** `src/powerquery/FDA_510k_Query.pq` (for reference)
* **Trigger:** Refreshed by VBA (`RefreshPowerQuery` function) during the `Workbook_Open` event.
* **Functionality:**
    * Calculates the start/end dates for the *previous* full calendar month.
    * Constructs a dynamic URL for the openFDA 510(k) endpoint, filtering by `decision_date`.
    * Includes a `limit=1000` parameter (assumes <1000 records/month).
    * Fetches JSON data via `Web.Contents`.
    * Handles potential API errors using `try...otherwise`.
    * Parses the JSON response and expands required fields (`k_number`, `applicant`, `deviceName`, dates, codes, etc.).
    * Renames fields to user-friendly names.
    * Performs basic type conversions (Date, Text).
    * Calculates `ProcTimeDays` (DecisionDate - DateReceived).
    * Constructs the `FDA_Link` URL.
    * Selects and reorders a base set of columns.
    * Sorts data by `DecisionDate` descending.
    * Loads the resulting table into the `CurrentMonthData` sheet, overwriting previous content.

### 3.3. Excel Sheets

* **`CurrentMonthData`:**
    * Target for the Power Query output.
    * Contains the main data table (`ListObject`).
    * VBA adds calculated columns directly to this table (scores, weights, recap, category).
    * Final formatting is applied here.
    * *(Optional)* Contains sheet event code (`Worksheet_Change`, `Worksheet_SelectionChange`) for detailed user activity logging.
* **`Weights`:**
    * Contains user-configurable Excel Tables (`ListObjects`):
        * `tblACWeights` (AC Code -> Weight)
        * `tblSTWeights` (Submission Type -> Weight)
        * `tblPCWeights` (Product Code -> Weight - Optional)
        * `tblKeywords` (High-value keywords for scoring)
        * `tblNFCosmeticKeywords` (Keywords triggering cosmetic negative factor)
        * `tblNFDiagnosticKeywords` (Keywords triggering diagnostic negative factor)
        * `tblTherapeuticKeywords` (Keywords potentially negating negative factors)
    * *(Optional)* May contain the `DebugMode` Named Range cell.
* **`CompanyCache`:**
    * Acts as a simple, persistent database for company summaries.
    * Columns: `CompanyName`, `RecapText`, `LastUpdated`.
    * VBA reads this into an in-memory dictionary (`dictCache`) on startup and saves the dictionary back to the sheet at the end of processing.
* **`RunLog` (Hidden - `xlSheetVeryHidden`):**
    * Created/managed by the `mod_Logger` module.
    * Stores a persistent history of processing runs, events, warnings, and errors.
    * Columns: `RunID`, `Timestamp`, `User`, `Step`, `Level`, `Message`, `Extra`.

### 3.4. VBA Modules

* **`ThisWorkbook`:**
    * Source: `src/vba/ThisWorkbook.cls`
    * Contains event handlers tied to the workbook itself.
    * `Workbook_Open()`: The main entry point. Sets up the application state, triggers `RefreshPowerQuery`, calls `ProcessMonthly510k`, handles cleanup.
    * *(Optional)* `Workbook_BeforeClose()`: Can be used to flush the log buffer (`FlushLogBuf`) and potentially trim the log (`TrimRunLog`).
* **`mod_510k_Processor`:**
    * Source: `src/vba/mod_510k_Processor.bas`
    * The main engine containing the core application logic.
    * **Constants:** Defines configuration (sheet names, API paths, scoring defaults, etc.).
    * **`ProcessMonthly510k()`:** Orchestrates the entire workflow after the initial refresh (Guard Check, Load Params, Process Loop, Write Back, Format, Archive Check, etc.).
    * **`Calculate510kScore()`:** Performs the core scoring calculation for a single record based on weights, keywords, NF/Synergy rules.
    * **Caching Functions:** `GetCompanyRecap()`, `GetCompanyRecapOpenAI()`, `LoadCompanyCache()`, `SaveCompanyCache()`. Handle in-memory cache logic, optional OpenAI calls, and persistence to the `CompanyCache` sheet.
    * **Loading Functions:** `LoadWeightsAndKeywords()`, `LoadTableToDict()`, `LoadTableToList()`. Read parameters from the `Weights` sheet tables into memory objects (Dictionaries, Collections).
    * **Column/Data Functions:** `AddScoreColumnsIfNeeded()`, `WriteResultsToArray()`, `GetColumnIndices()`. Manage table structure and data array manipulation.
    * **Formatting Functions:** `ReorganizeColumns()`, `FormatTableLook()`, `FormatCategoryColors()`, `ApplyNumberFormats()`, `CreateShortNamesAndComments()`, `FreezeHeaderAndKeyCols()`. Apply the final visual appearance.
    * **Archiving:** `ArchiveMonth()`. Handles copying the data sheet, renaming, and converting to static values.
    * **Helpers:** `GetWorksheets()`, `RefreshPowerQuery()` (Public), `IsMaintainerUser()`, `GetAPIKey()`, `SheetExists()`, JSON helpers, Safe data extractors.
* **`mod_Logger`:**
    * Source: `src/vba/mod_Logger.bas`
    * Provides a performance-optimized, buffered logging system.
    * **`LogEvt()`:** Adds log entries to an in-memory array.
    * **`FlushLogBuf()`:** Writes the entire buffer to the hidden `RunLog` sheet in one operation.
    * **`TrimRunLog()`:** *(Optional)* Deletes older entries from the `RunLog` sheet to prevent excessive growth.
    * Internal helpers (`InitLogger`, `EnsureLogSheet`, `DebugModeOn`).

## 4. Data Flow & Workflow

1.  **Workbook Open (`ThisWorkbook.Workbook_Open`):**
    * Application state optimized (ScreenUpdating Off, etc.).
    * `mod_510k_Processor.RefreshPowerQuery()` is called -> Power Query fetches data for the previous month -> Data loaded into `CurrentMonthData` table.
    * `mod_510k_Processor.ProcessMonthly510k()` is called.
2.  **Main Processing (`mod_510k_Processor.ProcessMonthly510k`):**
    * **Guard Check:** Determines if full processing should run (missing archive OR days 1-5 OR maintainer). If not, jumps to Archive Check.
    * **Setup:** Get sheet/table objects, ensure columns exist, map column indices.
    * **Load Parameters:** `LoadWeightsAndKeywords()` reads `Weights` sheet tables into memory dictionaries/collections. `LoadCompanyCache()` reads `CompanyCache` sheet into `dictCache`.
    * **Read Data:** `tblData.DataBodyRange.Value2` copied into `dataArr` variant array.
    * **Processing Loop:** Iterate through `dataArr`:
        * Call `Calculate510kScore()` using current row data and loaded parameters.
        * Call `GetCompanyRecap()` (checks `dictCache`, optionally calls `GetCompanyRecapOpenAI` if maintainer & needed).
        * Call `WriteResultsToArray()` to update the `dataArr` in memory with scores/recap.
    * **Write Back:** Write the entire modified `dataArr` back to the `CurrentMonthData` table.
    * **Format & Sort:** Apply number formats, sort the table (e.g., by `DecisionDate`).
    * **Save Cache:** Call `SaveCompanyCache()` to write `dictCache` back to the `CompanyCache` sheet.
    * **Final Layout:** Call `ReorganizeColumns()`, `FormatTableLook()`, `FormatCategoryColors()`, `CreateShortNamesAndComments()`, `FreezeHeaderAndKeyCols()`.
    * **(Jump Target `RunArchiveCheck`) Archive Check:** Calculate expected `archiveSheetName`. If `SheetExists()` returns false, call `ArchiveMonth()`.
    * **Logging:** `LogEvt` called at key stages. `FlushLogBuf` called before exit.
    * **Cleanup:** Restore application settings, release object variables.
3.  **(Optional) Workbook Close (`ThisWorkbook.Workbook_BeforeClose`):**
    * Final `FlushLogBuf`.
    * *(Optional)* Call `TrimRunLog`.

## 5. Configuration Points

* VBA Constants in `mod_510k_Processor` (Maintainer name, file paths, sheet/table names, default scores).
* Excel Tables on `Weights` sheet (Weights, Keywords).
* Named Range `DebugMode` on `Weights` sheet (optional).
* API Key stored in external file (`%APPDATA%\...`).

## 6. Dependencies

* Excel with VBA & Power Query.
* Windows OS (assumed).
* Internet connection.
* VBA References: `Microsoft Scripting Runtime`, `Microsoft XML, v6.0`.
* *(Optional)* OpenAI Account & API Key.

## 7. Error Handling

* Uses standard VBA `On Error GoTo <Label>` within major routines.
* Specific error handlers log details using `mod_Logger` (`LogEvt` with `lvlERROR` or `lvlWARN`).
* User-facing `MsgBox` used for critical errors preventing execution (missing sheets, columns, API key issues) or major warnings.
* `try...otherwise` used in Power Query for API call robustness.
* `On Error Resume Next` used sparingly for non-critical checks (e.g., finding optional columns, checking if comments exist before deleting).

## 8. Limitations & Considerations

* **API Rate Limits:** Assumes usage stays within openFDA and OpenAI API limits. No specific rate-limiting code implemented.
* **PQ Record Limit:** Assumes < 1000 records per month from FDA API. Pagination not implemented.
* **OpenAI Cost/Reliability:** OpenAI calls incur costs and depend on API availability. The basic JSON parsing is fragile.
* **Scalability:** Performance might degrade with extremely large numbers of keywords or cache entries, although array processing helps significantly.
* **Cross-Platform:** Developed/tested primarily on Windows; Mac compatibility may vary (especially regarding `Environ`, `WScript.Shell`, potential `MSXML` differences, file paths).