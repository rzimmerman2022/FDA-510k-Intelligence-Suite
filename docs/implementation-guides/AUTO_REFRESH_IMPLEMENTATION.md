# FDA 510(k) Intelligence Suite - Automatic Monthly Processing Implementation

## Quick Reference
- **Connection Name**: `Query - pgGet510kData` (If you rename your Power Query, update this constant in the Refresh510kConnection function)
- **Processing Rule**: 10th of month determines which month to archive
- **Target Month Logic**: Located in `DetermineReportingMonth()` function in ThisWorkbook.cls

## Overview

This document describes the implementation of an automatic monthly processing feature. The solution checks the current date when the workbook is opened and determines whether to process and archive the previous month's data based on the date.

## Implementation Details

### Date-Based Processing Rule

When the Excel workbook is opened, it follows these rules:
- If today's date is **on or after the 10th** of the current month, process the **previous month**
- If today's date is **before the 10th** of the current month, process the **month before previous**
- If the sheet for the target month already exists, no processing is performed

For example:
- If opened on May 8th, 2025: No auto-processing (before the 10th + March already exists)
- If opened on May 10th, 2025: Process April 2025 data (creates "Apr-2025" sheet)
- If opened on June 5th, 2025: Process April 2025 data (still prior to the 10th)
- If opened on June 10th, 2025: Process May 2025 data (creates "May-2025" sheet)

### Code Implementation

The solution was implemented by updating the `ThisWorkbook.cls` file with:

1. **Modified `Workbook_Open()` procedure** to:
   - Determine the appropriate reporting month based on current date
   - Check if the sheet for that month already exists
   - If needed, perform a synchronous Power Query refresh
   - Call the ProcessMonthly510k procedure to process the data

2. **Three new helper functions**:
   - `DetermineReportingMonth()`: Calculates which month to process based on current date
   - `SheetExists()`: Checks if a sheet with the given name already exists
   - `Refresh510kConnection()`: Refreshes the Power Query connection synchronously

### Technical Details

The code implements important technical safeguards:
- Uses synchronous refresh (BackgroundQuery = False) to avoid the Excel calculation race condition
- Handles year boundary cases (December → January) properly
- Includes multiple fallback mechanisms for finding and refreshing the connection
- Maintains proper UI state management (screen updating, calculation mode, etc.)
- Contains comprehensive error handling

## Benefits

This implementation:
1. Eliminates the need for manual monthly processing
2. Ensures data is processed on the appropriate schedule
3. Avoids duplicate processing when a month has already been archived
4. Uses synchronous refresh to prevent the calculation race condition that was causing errors

## Disabling/Rollback Instructions

If you need to disable the automatic processing temporarily or permanently:

### Option 1: Comment Out Code (Temporary Disable)
1. Open the VBA Editor (Alt+F11)
2. Navigate to ThisWorkbook in the Project Explorer
3. Comment out the code blocks in Workbook_Open that perform date checking and processing:
   ```vba
   ' Comment out this section to disable auto-processing
   ' tgtMonthFirst = DetermineReportingMonth(Date)
   ' tgtSheetName = Format(tgtMonthFirst, "mmm-yyyy")
   ' ...etc...
   ```

### Option 2: Complete Rollback
If you need to completely remove this feature:
1. Open the attached backup copy of ThisWorkbook.cls
2. Replace the current version with the backup version
3. Alternatively, you can undo all the changes by:
   - Removing the 3 helper functions (DetermineReportingMonth, SheetExists, Refresh510kConnection)
   - Restoring the previous version of Workbook_Open

### Option 3: Configuration Flag
For future enhancement: Consider adding a configuration setting in mod_Config.bas like:
```vba
Public Const AUTO_PROCESSING_ENABLED As Boolean = True ' Set to False to disable
```
And then modify Workbook_Open to check this flag before proceeding with auto-processing.

## Handling Edge Cases

### Year Boundaries
The implementation properly handles year boundary cases:
- When processing January data in early March (before the 10th), the code correctly handles the December→January transition
- DateSerial automatically wraps negative month values to the previous year

### Long Gaps in Use
The current implementation processes only the month determined by the current date rule. If the workbook hasn't been opened for several months, it will only process the most recently due month, not all skipped months.

If back-filling multiple missed months is required, a separate "catch-up" routine could be developed that iterates through missing months.

## Validation

The implementation was validated against the test cases mentioned above to ensure it properly:
- Identifies the correct month to process based on the current date
- Skips processing when the target sheet already exists
- Performs a synchronous refresh of the Power Query data
- Successfully calls ProcessMonthly510k to process and archive the data

## Future Enhancements

1. Move helper functions to a dedicated module (e.g., mod_UtilsMonthly) to keep ThisWorkbook.cls clean
2. Modify ProcessMonthly510k to accept a target month parameter to avoid redundant date calculations
3. Add a configuration flag in mod_Config.bas to easily enable/disable auto-processing
4. Add capability to process multiple missing months if the workbook hasn't been used for a long time
