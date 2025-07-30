# FDA 510(k) Intelligence Suite - User Guide

## Introduction

Welcome! This Excel workbook automatically finds, scores, and presents recent FDA 510(k) clearances relevant to our interests. Its goal is to help you quickly identify the most important leads from the previous month.

## Getting Started

1.  **Open the File:** Navigate to `assets/excel-workbooks/` and double-click the main `.xlsm` workbook file.
2.  **Enable Content (IMPORTANT!):** When Excel opens, you will likely see yellow security bars near the top. You **must** click **"Enable Content"** and/or **"Enable Macros"**. If you don't, the automated process will not run.
    * *If you don't see these buttons but the tool doesn't seem to work, check Excel's Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings & Message Bar settings) or contact the maintainer.*
3.  **Automatic Processing:** Once content is enabled, the workbook automatically does the following:
    * **Refreshes Data:** It connects to the FDA database to get the latest data for the *previous* full month. You'll see messages in the bottom status bar like "Refreshing FDA data...".
    * **Scores & Formats:** It then calculates scores, generates company summaries (where possible), and formats the main sheet. You'll see status bar messages like "Processing leads...". This step might be skipped if it's late in the month (after the 5th) - this is normal.
    * **Archives (Monthly):** On the first run within the first few days of a new month, it automatically creates a backup sheet (e.g., "Apr-2025") of the previous month's final data.
4.  **Ready:** When the status bar shows "Workbook ready" or "Refresh complete", the process is finished.

## Understanding the `CurrentMonthData` Sheet

This is the main sheet you'll use. It displays the processed data for the *previous* month.

* **Key Columns (Left Side):**
    * `K_Number`, `DecisionDate`, `Applicant`, `DeviceName`, `Contact`: Core information about the clearance.
    * `CompanyRecap`: Provides a quick summary of the applicant company. It might say "Needs Research" (meaning no summary is available in the local cache) or show a 1-2 sentence AI-generated summary (if enabled by the maintainer).
    * `Score_Percent`: The calculated relevance score for this lead (higher is generally more relevant).
    * `Category`: A quick category based on the score:
        * <span style="background-color:#C6EFCE; color:black;">High</span> (Greenish): Potentially very relevant.
        * <span style="background-color:#FFEB9C; color:black;">Moderate</span> (Yellowish): Worth a look.
        * <span style="background-color:#FFDDCC; color:gray;">Low</span> (Orangish): Less likely relevant.
        * <span style="background-color:#F2F2F2; color:gray;">Almost None</span> (Gray): Probably not relevant.
        * <span style="background-color:#FFC7CE; color:red;">Error</span> (Reddish): Indicates an error during scoring for this row.
    * `FDA_Link`: A direct web link to the official FDA record for this 510(k).
* **Device Name Comments:** If a `DeviceName` is very long, it might be shortened in the cell and end with "...". **Hover your mouse cursor over the cell** to see the full, original device name appear in a comment box.
* **Thick Border:** There's a thicker line after the `FDA_Link` column. Columns to the **right** of this border contain the detailed breakdown of how the score was calculated (individual weights, etc.) and other raw data. You typically don't need to focus on these unless investigating a specific score.

## Interacting with the Data

* **Sorting & Filtering:** This sheet uses a standard Excel Table. You can use the **dropdown arrows in the header row** to sort columns (e.g., sort by `Score_Percent` descending to see highest scores first) or filter them (e.g., filter `Category` to show only "High" and "Moderate").
* **Copying:** You can select cells or rows and copy/paste them into emails or other documents as needed.

## Archive Sheets

* Sheets named with a month and year (e.g., `Apr-2025`, `Mar-2025`) are **historical archives**.
* They contain a static snapshot of the data as it was processed for that month.
* They do not update automatically. Use the `CurrentMonthData` sheet for the most recently processed data.

## Simple Troubleshooting

* **Data Not Updating?:** Ensure you clicked "Enable Content". Check your internet connection. If problems persist, contact the maintainer.
* **Error Messages:** If you see a pop-up error message during opening or processing, please note down the message text and contact the maintainer.
* **Need More Detail?:** A hidden `RunLog` sheet exists that tracks the tool's operation. Contact the maintainer if detailed logs are needed to investigate an issue.

## Contact

For any issues, questions, or suggestions regarding this tool, please contact **[Maintainer's Name - e.g., Ryan Zimmerman]**.