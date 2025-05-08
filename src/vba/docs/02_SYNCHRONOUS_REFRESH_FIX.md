# FDA 510(k) Intelligence Suite - Synchronous Refresh Fix

## Problem

The application was experiencing errors when refreshing Power Query data with BackgroundQuery = True (asynchronous mode) combined with Automatic Calculation mode. This created a race condition where:

1. The asynchronous query would start in a background thread
2. Excel's calculation engine would attempt to recalculate while the query was still running
3. Both engines would compete for the workbook lock
4. This resulted in the error: "Microsoft Excel is refreshing some data. Please try again later."

## Diagnostic Results

From the test logs, we can see a systematic pattern:

| Configuration | Result | Explanation |
|---------------|--------|-------------|
| BackgroundQuery = False + Calculation = Manual | ✅ Success | Synchronous refresh (waits for completion) with no auto-calculation |
| BackgroundQuery = False + Calculation = Automatic | ✅ Success | Synchronous refresh completes before auto-calculation starts |
| BackgroundQuery = True + Calculation = Manual | ✅ Success | Async refresh works because manual calculation prevents recalc conflicts |
| BackgroundQuery = True + Calculation = Automatic | ❌ Failure | The problematic case: race condition between async refresh and auto-calculation |

The test diagnostics also revealed that the refresh works successfully when triggered from the Timer event context, confirming this is a thread synchronization issue.

## Solution Implemented

We've implemented a robust fix that:

1. Always uses synchronous refresh (BackgroundQuery = False) in the Refresh510kConnection() function
2. Directly targets the WorkbookConnection when possible, falling back to the traditional QueryTable method if needed
3. Temporarily switches Excel to Manual Calculation mode during the workbook open procedure

This approach ensures that:
- The Power Query refresh completes fully before any calculations occur
- No race condition can occur between the refresh and calculation processes
- The system gracefully handles both direct connection refresh and fallback methods

## Technical Implementation

The implementation adds a new Refresh510kConnection() function to ThisWorkbook that:

```vba
Public Function Refresh510kConnection() As Boolean
    ' *** IMPORTANT: If you rename your Power Query in the QueryEditor, update this constant ***
    Const CN_NAME As String = "Query - pgGet510kData" ' Default connection name pattern
    
    ' Try the direct connection approach first
    Set cn = ThisWorkbook.Connections(CN_NAME)
    If Not cn Is Nothing Then
        ' Set to synchronous mode and refresh
        cn.OLEDBConnection.BackgroundQuery = False   ' <-- Forces synchronous refresh
        cn.Refresh
        Refresh510kConnection = True
    Else
        ' Fall back to mod_DataIO if needed
        Refresh510kConnection = mod_DataIO.RefreshPowerQuery(tblData)
    End If
End Function
```

> **Important Note:** The default connection name pattern is `Query - pgGet510kData`. If you rename your Power Query in the Power Query Editor, make sure to update the `CN_NAME` constant in the Refresh510kConnection function accordingly.

## Benefits

1. **Reliability**: Eliminates the race condition error completely
2. **Simplicity**: Uses a straightforward approach (synchronous refresh) rather than complex event-handling
3. **Performance**: While synchronous refresh blocks the UI, it's only done once per month during the auto-processing cycle
4. **Compatibility**: Works with all Excel versions and configurations

## Alternatives Considered

We considered other approaches including:

1. **Event-based asynchronous refresh**: More complex, requires sinking events and risks event timing issues
2. **Timed retry loops with DoEvents**: Less reliable, could still hit race conditions
3. **Multiple calculation mode toggles**: Less preferred as it adds complexity without benefit over using synchronous refresh

The implemented solution provides the best balance of simplicity, reliability and maintainability.
