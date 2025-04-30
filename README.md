# FDA 510(k) Intelligence Suite (Excel VBA Tool)

**Automated fetching, scoring, caching, and archiving of FDA 510(k) clearance data directly within Microsoft Excel.**

This tool streamlines the process of identifying and prioritizing relevant FDA 510(k) clearances based on configurable scoring models, saving significant manual effort.

## Features

*   **Automated Data Fetching:** Uses Power Query to dynamically pull the previous full month's 510(k) data from the official openFDA API.
*   **Configurable Scoring:** Implements a weighted scoring model based on factors like Advisory Committee (AC), Product Code (PC), Keywords (KW), Submission Type (ST), Processing Time (PT), and Geography/Location (GL). Weights and keywords are managed via simple Excel tables.
*   **Negative Factors & Synergy:** Incorporates rules for negative factors (e.g., cosmetic, diagnostic w/o therapeutic) and synergy bonuses based on keyword lists and specific criteria.
*   **Company Recaps (Optional AI):**
    *   Includes a local cache (`CompanyCache` sheet) for company summaries.
    *   *(Maintainer Only)* Optionally leverages the OpenAI API (GPT-3.5-Turbo or similar) to generate concise company summaries if not found in the cache (requires API key).
*   **Automated Workflow:** Runs automatically on workbook open (`Workbook_Open`):
    *   Refreshes Power Query data.
    *   Conditionally runs scoring/formatting (based on date guard/maintainer status).
    *   Checks if the previous month's archive exists and creates it if missing.
*   **Dynamic Formatting:** Applies conditional formatting (colors based on score category), specific table styles, column widths, and handles potential duplicate columns. Shortens long device names with full text available in comments. *(Note: Pane freezing is currently disabled based on recent feedback)*.
*   **Robust Logging:** Includes a detailed, buffered logging system (`mod_Logger`) writing to a hidden `RunLog` sheet for diagnostics and history tracking with minimal performance impact. Also includes a secondary, more verbose tracing system (`mod_DebugTraceHelpers`) writing to a `DebugTrace` sheet, controlled by constants in that module.
*   **Maintainer Overrides:** Allows a designated maintainer (set via `MAINTAINER_USERNAME` in `mod_Config`) to bypass date guards and enable optional features like OpenAI calls and detailed logging (via `mod_Logger`'s `DebugModeOn` function, which checks maintainer status and the `DebugMode` named range).
*   **Modular Design:** Code is organized into distinct VBA modules (e.g., `mod_DataIO`, `mod_Score`, `mod_Format`, `mod_Config`) for better maintainability.
*   **Development Utilities:** Includes a `ModuleManager` module to help export/import VBA code for version control.

## Screenshot

*(This shows the typical layout of the CurrentMonthData sheet after processing)*
![CurrentMonthData Sheet Layout](docs/images/image_298e21.jpg)
*(You might need to adjust the path if you move the image, or embed it directly if your Markdown host supports it)*

## Requirements

*   **Microsoft Excel:** Version supporting Power Query and VBA Macros (.xlsm format). Tested primarily on Windows.
*   **Windows OS:** Recommended due to dependencies like `Environ("USERNAME")`, `WScript.Shell`, `MSXML2.ServerXMLHTTP.6.0`.
*   **Internet Connection:** Required for Power Query API calls and optional OpenAI calls.
*   **OpenAI API Key (Optional):** Only required for the *maintainer* if automated company summaries via OpenAI are desired.
*   **VBA References (Check VBE -> Tools -> References):**
    *   Microsoft Scripting Runtime
    *   Microsoft Visual Basic for Applications Extensibility 5.3 (for ModuleManager)
    *   Microsoft XML, v6.0 (for OpenAI calls)

## Setup & Installation

1.  **Clone Repository:** `git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git`
2.  **Open Workbook:** Open the `FDA-510k-Intelligence-Suite.xlsm` file (or your specific `.xlsm` file name).
3.  **Enable Content:** When prompted, **Enable Macros** and **Enable Content** (for Power Query data connections).
4.  **Configure Maintainer:**
    *   Press `Alt+F11` to open the VBA Editor.
    *   In the Project Explorer, navigate to `Modules` -> `mod_Config`.
    *   Find the constant `MAINTAINER_USERNAME` near the top.
    *   **IMPORTANT:** Change `"YourWindowsUsername"` to your exact Windows login username. This enables maintainer-specific features like OpenAI calls and bypassing certain processing guards.
5.  **Configure Weights & Keywords:**
    *   Go to the `Weights` worksheet in Excel.
    *   Populate the tables (`tblACWeights`, `tblSTWeights`, `tblPCWeights`, `tblKeywords`, `tblNFCosmeticKeywords`, `tblNFDiagnosticKeywords`, `tblTherapeuticKeywords`) with your desired codes, keywords, and scoring weights.
6.  **Configure OpenAI API Key (Maintainer Only - Optional):**
    *   Create the folder path specified by the `API_KEY_FILE_PATH` constant in `mod_Config` (default: `%APPDATA%\510k_Tool\`). You might need to show hidden folders to see `AppData`.
    *   Inside that folder, create a plain text file named `openai_key.txt`.
    *   Open the file and paste **only** your OpenAI API key into it. Save and close.
    *   *Security Note:* This method keeps the key out of the code and repository, but relies on filesystem access control.
7.  **Configure Detailed Logging (Maintainer Only - Optional):**
    *   **Method 1 (Named Range for mod_Logger):**
        *   Go to the `Weights` sheet (or any sheet).
        *   Select an unused cell.
        *   Go to the **Name Box** (left of the formula bar), type `DebugMode`, and press Enter.
        *   Enter `TRUE` in the named cell to enable detailed logging via `mod_Logger` (only effective if you are the `MAINTAINER_USERNAME`), or `FALSE` to disable it.
    *   **Method 2 (Environment Variable for mod_Logger):** Set a system environment variable `TRACE_ALL_USERS` to `1` to force detailed logging for all users via `mod_Logger` (overrides named range).
    *   **Method 3 (Constants for mod_DebugTraceHelpers):** For the separate `mod_DebugTraceHelpers` system, edit the `TRACE_ENABLED` and `TRACE_LEVEL` constants directly within that module in the VBE.
8.  **Save:** Save the workbook.

## Usage

1.  **Open the Workbook:** Double-click the `.xlsm` file.
2.  **Automatic Refresh:** On open, Power Query will automatically attempt to refresh data for the *previous* full month. The status bar will indicate progress. *(Note: Refresh may be disabled by default after a successful run; the code now re-enables it before attempting)*.
3.  **Automatic Processing:**
    *   The VBA code (`ProcessMonthly510k`) will then run.
    *   It checks if full processing (scoring, formatting) should occur based on whether the archive sheet for the previous month already exists, the current day (first 5 days), or if you are the maintainer.
    *   If processing runs, it calculates scores, fetches/caches recaps, writes results, applies formatting, and saves the cache.
    *   It then checks if the archive sheet for the previous month exists. If not, it creates it by copying the current data sheet and converting it to static values.
4.  **Review Data:** Examine the `CurrentMonthData` sheet for the scored and formatted leads. Hover over shortened `DeviceName` entries (if implemented) to see the full text.
5.  **Review Log (Optional):** If troubleshooting, unhide the `RunLog` sheet (requires VBA: `ThisWorkbook.Sheets("RunLog").Visible = xlSheetVisible`) to view detailed run history and errors from `mod_Logger`. Hide it again afterwards (`.Visible = xlSheetVeryHidden`).
6.  **Review Trace (Optional):** Unhide the `DebugTrace` sheet to view verbose trace information from `mod_DebugTraceHelpers` (if enabled).

## Architecture

The tool utilizes a combination of Excel features and VBA modules:
*   **Power Query:** For robust data fetching and initial transformation from the openFDA API.
*   **VBA Modules:**
    *   `ThisWorkbook`: Handles workbook events (`Open`, `BeforeClose`).
    *   `mod_510k_Processor`: Orchestrates the main processing workflow.
    *   `mod_Config`: Centralizes global constants and settings.
    *   `mod_DataIO`: Handles Power Query refresh, sheet checks, array writing, connection cleanup.
    *   `mod_Schema`: Manages column mapping and safe data access by name.
    *   `mod_Weights`: Loads and provides access to scoring weights and keywords.
    *   `mod_Score`: Contains the core scoring calculation logic.
    *   `mod_Cache`: Manages company recap caching (memory and sheet) and optional OpenAI interaction.
    *   `mod_Format`: Applies all visual formatting, column reordering, sorting, etc.
    *   `mod_Archive`: Handles the creation of monthly archive sheets.
    *   `mod_Logger`: Provides buffered logging to the `RunLog` sheet.
    *   `mod_DebugTraceHelpers`: Provides conditional, verbose tracing to the `DebugTrace` sheet.
    *   `mod_Debug`: Contains older/simpler debugging utilities.
    *   `ModuleManager`: Provides utilities for exporting/importing VBA code.
*   **Excel Tables:** For managing scoring parameters (weights, keywords) on the `Weights` sheet.
*   **Excel Sheets:** For data display (`CurrentMonthData`), parameters (`Weights`), caching (`CompanyCache`), logging (`RunLog`), and tracing (`DebugTrace`).

See [ARCHITECTURE.md](docs/ARCHITECTURE.md) for a more detailed breakdown.
See [AI_DEVELOPMENT_GUIDE.md](docs/AI_DEVELOPMENT_GUIDE.md) for guidelines on coding and commenting practices when using AI assistance.

## Configuration

Key configuration points are centralized in the VBA code within the **`mod_Config`** module:
*   `MAINTAINER_USERNAME`: Your Windows login name. **MUST BE SET.**
*   `API_KEY_FILE_PATH`: Path to the optional OpenAI key file. **VERIFY PATH.**
*   Sheet Names (`DATA_SHEET_NAME`, `WEIGHTS_SHEET_NAME`, etc.).
*   Scoring Defaults and Rule Constants (`DEFAULT_AC_WEIGHT`, `NF_COSMETIC`, etc.). **These should be reviewed against your specific scoring model.**
*   OpenAI settings (`OPENAI_API_URL`, `OPENAI_MODEL`, etc.).
*   UI/Formatting constants (`RECAP_MAX_LEN`, etc.).
*   `VERSION_INFO`.

Parameter tables on the `Weights` sheet also control scoring behavior.
Tracing behavior is controlled by constants in `mod_DebugTraceHelpers` and the `DebugMode` named range / environment variable for `mod_Logger`.

## Contributing

*(Optional: Add guidelines if others might contribute. Link to CONTRIBUTING.md if created).*

## License

*(Optional: Specify the license under which this project is shared. Link to LICENSE file if created).*
