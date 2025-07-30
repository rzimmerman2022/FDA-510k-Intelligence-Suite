# Power Query Refresh Fix - Implementation Guide

This document provides step-by-step instructions for implementing the timer-based refresh solution to resolve the Error 1004 issue in the FDA-510k-Intelligence-Suite.

## Prerequisites

Before implementing the solution, ensure you have:

1. A backup of the workbook
2. The latest version of the VBA code
3. Access to modify VBA modules

## Implementation Steps

### Step 1: Modify the mod_DataIO Module

1. Open the VBA editor (Alt+F11)
2. Navigate to `mod_DataIO` in the Project Explorer
3. Add the module-level variables at the top of the module, after `Option Explicit`:

```vba
' --- Module-level variables for timer-based refresh ---
Private mTargetTableToRefresh As ListObject ' Holds the table being refreshed via timer
Private mRefreshTimer As Double             ' Holds the timer ID
Private mRefreshSuccess As Boolean          ' Holds the result of the async refresh operation
Private mRefreshAttemptCount As Integer     ' Counter for retries
Private mRefreshErrorMessage As String      ' Holds any error message
Private mRefreshWaitingForCompletion As Boolean ' Flag to indicate refresh is pending
Private Const MAX_REFRESH_ATTEMPTS As Integer = 2 ' Max number of retry attempts
```

### Step 2: Replace the RefreshPowerQuery Function

1. Find the existing `RefreshPowerQuery` function
2. Replace it completely with the new timer-based version from the `POWER_QUERY_REFRESH_FIX.md` document

### Step 3: Add New Helper Functions

Add these new functions to the module:

1. `ExecuteTimerRefresh` - The function called by the timer
2. `CancelRefresh` - Optional function to cancel pending refreshes
3. `DiagnoseTableState` - Diagnostics function for table state
4. `DiagnoseConnectionState` - Diagnostics function for connection state

All these functions can be copied from the `POWER_QUERY_REFRESH_FIX.md` document.

### Step 4: Test the Implementation

1. Compile the VBA project:
   - In the VBA editor, go to Debug > Compile VBAProject
   - Fix any syntax errors that may appear

2. Save and close the workbook

3. Reopen the workbook
   - When prompted to refresh data, click "Yes"
   - Monitor the progress in the status bar

4. Check for successful refresh:
   - Look at the RunLog sheet to see detailed logs
   - Verify the data has been refreshed with the latest information

### Step 5: Test the Monthly Process

1. Run the monthly process routine:
   ```vba
   Application.Run "ProcessMonthly510k"
   ```

2. Observe if the refresh completes successfully as part of the workflow

## Troubleshooting

### Compilation Errors

If you encounter compilation errors:

1. Check for missing references:
   - In the VBA editor, go to Tools > References
   - Ensure all required references are checked

2. Verify variable declarations:
   - Make sure all variables used in the new code are properly declared
   - Look for typos in variable names

### Runtime Errors

If the refresh still fails with errors:

1. Check the logs:
   - Look at the RunLog sheet for detailed error messages
   - Pay attention to the `Connection.Refreshing` and `OLEDBConnection` status logs

2. Verify connection names:
   - The code uses "pgGet510kData" and "Query - pgGet510kData" as default connection names
   - You may need to modify these to match your actual connection names if different

3. Debug mode:
   - Set a breakpoint at the beginning of `RefreshPowerQuery`
   - Step through the execution (F8) to identify where the issue occurs

### Other Issues

1. **Timer Not Firing**: If the timer doesn't seem to fire:
   - Check if Application.EnableEvents is set to True
   - Verify the ThisWorkbook.Name property is correct in the OnTime call

2. **Long Refresh Times**: If the refresh takes too long:
   - Consider increasing the timeout period (currently 60 seconds)
   - Add a progress indicator to keep the user informed

## Notes for Advanced Users

1. **Customizing Retry Logic**:
   - You can adjust the `MAX_REFRESH_ATTEMPTS` constant to increase retries
   - You can modify the retry delay (currently 2 seconds) in `ExecuteTimerRefresh`

2. **Extending Diagnostics**:
   - The diagnostic functions can be extended to log additional information
   - Consider adding more detailed checks for complex environments

3. **Error Handling**:
   - The solution includes robust error handling with detailed logs
   - Review the logs carefully when troubleshooting issues

## Best Practices

1. Always maintain a backup of the original code
2. Test the solution thoroughly before deploying to production
3. Document any customizations you make to the solution
4. Monitor logs regularly to identify potential issues early

---

For a complete technical explanation of the solution, refer to the `POWER_QUERY_REFRESH_FIX.md` document.
