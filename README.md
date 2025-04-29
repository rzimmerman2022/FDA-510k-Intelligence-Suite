# FDA 510(k) Intelligence Suite (Excel VBA Tool)

**Automated fetching, scoring, caching, and archiving of FDA 510(k) clearance data directly within Microsoft Excel.**

This tool streamlines the process of identifying and prioritizing relevant FDA 510(k) clearances based on configurable scoring models, saving significant manual effort.

## Features

* **Automated Data Fetching:** Uses Power Query to dynamically pull the previous full month's 510(k) data from the official openFDA API.
* **Configurable Scoring:** Implements a weighted scoring model based on factors like Advisory Committee (AC), Product Code (PC), Keywords (KW), Submission Type (ST), Processing Time (PT), and Geography/Location (GL). Weights and keywords are managed via simple Excel tables.
* **Negative Factors & Synergy:** Incorporates rules for negative factors (e.g., cosmetic, diagnostic w/o therapeutic) and synergy bonuses based on keyword lists and specific criteria.
* **Company Recaps (Optional AI):**
    * Includes a local cache (`CompanyCache` sheet) for company summaries.
    * *(Maintainer Only)* Optionally leverages the OpenAI API (GPT-3.5-Turbo or similar) to generate concise company summaries if not found in the cache (requires API key).
* **Automated Workflow:** Runs automatically on workbook open (`Workbook_Open`):
    * Refreshes Power Query data.
    * Conditionally runs scoring/formatting (based on date guard/maintainer status).
    * Checks if the previous month's archive exists and creates it if missing.
* **Dynamic Formatting:** Applies conditional formatting (colors based on score category), specific table styles, column widths, frozen panes, and shortens long device names with full text available in comments.
* **Robust Logging:** Includes a detailed, buffered logging system (`mod_Logger`) writing to a hidden `RunLog` sheet for diagnostics and history tracking with minimal performance impact.
* **Maintainer Overrides:** Allows a designated maintainer (set via `MAINTAINER_USERNAME`) to bypass date guards and enable optional features like OpenAI calls and detailed logging (`DebugMode` named range).

## Screenshot

*(This shows the typical layout of the CurrentMonthData sheet after processing)*
![CurrentMonthData Sheet Layout](docs/images/image_298e21.jpg)
*(You might need to adjust the path if you move the image, or embed it directly if your Markdown host supports it)*

## Requirements

* **Microsoft Excel:** Version supporting Power Query and VBA Macros (.xlsm format). Tested primarily on Windows.
* **Windows OS:** Recommended due to dependencies like `Environ("USERNAME")`, `WScript.Shell`, `MSXML2.ServerXMLHTTP.6.0`.
* **Internet Connection:** Required for Power Query API calls and optional OpenAI calls.
* **OpenAI API Key (Optional):** Only required for the *maintainer* if automated company summaries via OpenAI are desired.

## Setup & Installation

1.  **Clone Repository:** `git clone https://github.com/rzimmerman2022/FDA-510k-Intelligence-Suite.git`
2.  **Open Workbook:** Open the `FDA-510k-Intelligence-Suite.xlsm` file (or your specific `.xlsm` file name).
3.  **Enable Content:** When prompted, **Enable Macros** and **Enable Content** (for Power Query data connections).
4.  **Configure Maintainer:**
    * Press `Alt+F11` to open the VBA Editor.
    * In the Project Explorer, navigate to `Modules` -> `mod_510k_Processor`.
    * Find the constant `MAINTAINER_USERNAME` near the top.
    * **IMPORTANT:** Change `"YourWindowsUsername"` to your exact Windows login username.
5.  **Configure Weights & Keywords:**
    * Go to the `Weights` worksheet in Excel.
    * Populate the tables (`tblACWeights`, `tblSTWeights`, `tblPCWeights`, `tblKeywords`, `tblNFCosmeticKeywords`, `tblNFDiagnosticKeywords`, `tblTherapeuticKeywords`) with your desired codes, keywords, and scoring weights.
6.  **Configure OpenAI API Key (Maintainer Only - Optional):**
    * Create the folder path: `%APPDATA%\510k_Tool\` (e.g., `C:\Users\YourUser\AppData\Roaming\510k_Tool\`). You might need to show hidden folders to see `AppData`.
    * Inside that folder, create a plain text file named `openai_key.txt`.
    * Open the file and paste **only** your OpenAI API key into it. Save and close.
    * *Security Note:* This method keeps the key out of the code and repository, but relies on filesystem access control.
7.  **Configure Debug Mode (Maintainer Only - Optional):**
    * Go to the `Weights` sheet.
    * Select an unused cell (e.g., A10).
    * Go to the **Name Box** (left of the formula bar), type `DebugMode`, and press Enter.
    * Enter `FALSE` in the named cell to disable detailed logging, or `TRUE` to enable it (only works if you are also the `MAINTAINER_USERNAME`).
8.  **Save:** Save the workbook.

## Usage

1.  **Open the Workbook:** Double-click the `.xlsm` file.
2.  **Automatic Refresh:** On open, Power Query will automatically refresh data for the *previous* full month. The status bar will indicate progress.
3.  **Automatic Processing:**
    * The VBA code (`ProcessMonthly510k`) will then run.
    * It checks if full processing (scoring, formatting) should occur based on whether the archive sheet for the previous month already exists, the current day (first 5 days), or if you are the maintainer.
    * If processing runs, it calculates scores, fetches/caches recaps, writes results, applies formatting, and saves the cache.
    * It then checks if the archive sheet for the previous month exists. If not, it creates it by copying the current data sheet and converting it to static values.
4.  **Review Data:** Examine the `CurrentMonthData` sheet for the scored and formatted leads. Hover over shortened `DeviceName` entries to see the full text.
5.  **Review Log (Optional):** If troubleshooting, unhide the `RunLog` sheet (requires VBA: `ThisWorkbook.Sheets("RunLog").Visible = xlSheetVisible`) to view detailed run history and errors. Hide it again afterwards (`.Visible = xlSheetVeryHidden`).

## Architecture

The tool utilizes a combination of Excel features:
* **Power Query:** For robust data fetching and initial transformation.
* **VBA:** For orchestration, custom scoring logic, API interactions (OpenAI), caching, formatting, and archiving.
* **Excel Tables:** For managing scoring parameters (weights, keywords).
* **Excel Sheets:** For data display (`CurrentMonthData`), parameters (`Weights`), caching (`CompanyCache`), and logging (`RunLog`).

See [ARCHITECTURE.md](docs/ARCHITECTURE.md) for a more detailed breakdown.

## Configuration

Key configuration points within the VBA code (`mod_510k_Processor` module constants):
* `MAINTAINER_USERNAME`: Your Windows login name.
* `API_KEY_FILE_PATH`: Path to the optional OpenAI key file.
* `PQ_CONNECTION_NAME`: Name of the Power Query connection.
* Sheet Names (`DATA_SHEET_NAME`, etc.).
* Scoring Defaults and Rule Constants (`DEFAULT_AC_WEIGHT`, `NF_COSMETIC`, etc.). **These should be reviewed against your specific scoring model.**

Parameter tables on the `Weights` sheet also control scoring behavior.

## Contributing

*(Optional: Add guidelines if others might contribute. Link to CONTRIBUTING.md if created).*

## License

*(Optional: Specify the license under which this project is shared. Link to LICENSE file if created).*