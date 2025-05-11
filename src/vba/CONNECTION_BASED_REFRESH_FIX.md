# Power Query Connection-Based Refresh Fix

## Problem Fixed

The FDA 510k Intelligence Suite was experiencing Error #1004 ("Application-defined or object-defined error") during the Power Query refresh process. After investigation, we identified that the error occurred because:

1. The workbook uses a "Connection-only" Power Query that's loaded to a table
2. The code was trying to refresh through the QueryTable interface (legacy approach)
3. Modern Power Query connections must be refreshed through the WorkbookConnection object

## Implementation Details

### Two-Step Fix Applied:

1. **Updated Refresh Path**
   - Modified `RefreshPowerQuery` in `mod_DataIO.bas` to first attempt refresh via WorkbookConnection object
   - Added fallback to QueryTable.Refresh only if WorkbookConnection cannot be found
   - This handles "Connection-only" queries that produce tables via "Load To Table"

2. **Improved Connection Handling**
   - Added automatic cleanup of duplicate connections before each refresh attempt 
   - Removed setting `EnableRefresh = False` after refresh, since toggling this can invalidate connections
   - Simplified error messages and improved logging for better diagnostics

### Related Improvements

The duplicate connections issue is now addressed at two levels:

1. **Prevention**: The recent update to `mod_Archive.bas` (May 2025) creates values-only copies for archives, which prevents duplicate connection creation.

2. **Cleanup**: The `CleanupDuplicateConnections` function is now called automatically before each refresh operation to remove any duplicate connections that might exist.

## Verification

To verify the fix is working properly:

1. **Manual Test**
   - Open the workbook and select Data ▸ Refresh All
   - Verify it completes without error

2. **VBA Test**
   - In the Immediate window (Ctrl+G), run:
     ```vba
     ?mod_DataIO.RefreshPowerQuery(Sheets("CurrentMonthData").ListObjects(1))
     ```
   - Verify it returns `True`

3. **Full Pipeline Test**
   - Run `ProcessMonthly510k` and verify no 1004 errors appear in the log

## Technical Background

Modern Power Query tables in Excel use a multi-layered architecture:

1. **WorkbookConnection**: The core object managing the query, data model connection, and refresh pipeline
2. **ListObject**: The table UI component displayed on the worksheet
3. **QueryTable**: A legacy compatibility object that acts as a shim between ListObjects and connections

When using "Connection-only" mode with "Load To Table", Excel creates all three objects, but the QueryTable is just a thin wrapper that delegates refresh operations to the WorkbookConnection. Attempting direct refresh via QueryTable leads to Error #1004.

Additionally, copying sheets with PQ-connected tables creates duplicate connections with numerical suffixes (e.g., "pgGet510kData (2)"), which causes another class of errors. The new values-only archiving approach prevents this issue.
