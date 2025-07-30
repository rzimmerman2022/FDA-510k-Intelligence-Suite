# Connection Name Fix

## Issue Overview

The Power Query refresh process was failing with Error #1004 because the code was attempting to refresh a QueryTable associated with a Power Query ListObject, which is not the recommended approach. The root cause was specifically a name mismatch - the code was looking for a WorkbookConnection with the exact name matching the table name (e.g., "pgGet510kData"), but the actual connection was prefixed with "Query - " (e.g., "Query - pgGet510kData").

## Error Symptoms

- RunLog showed no "Found WorkbookConnection" message, indicating the code path to grab the WorkbookConnection wasn't running
- Error #1004 immediately followed, with the message "QueryTable.Refresh can't refresh a PQ list-object"
- The code was falling back to the legacy QueryTable approach which is known to fail with Error #1004 for Power Query tables

## Verification Test Results

Testing in the immediate window confirmed the connection existed with the name format "Query - pgGet510kData":

```vba
For Each c In ThisWorkbook.Connections: ? c.Name: Next
Query - pgGet510kData

? ThisWorkbook.Connections("Query - pgGet510kData").RefreshWithRefreshAll
True
```

The `True` result from the second command proved the connection itself refreshes fine when directly referenced by its correct name.

## Solution Implemented

The fix involved making two major changes:

1. **Enhanced `RefreshPowerQuery` Function** - Completely redesigned the function to first search for the WorkbookConnection using multiple approaches:
   - Try various exact name patterns (baseName, "Query - " + baseName, "Connection " + baseName)
   - If no exact match, perform a loose search for any connection containing the baseName
   - Only fall back to QueryTable approach if absolutely no matching connection can be found

2. **Updated Function Parameters** - Added parameters to make the function more flexible:
   - `baseName` parameter allows specifying the base name to search for, defaulting to the table name if not specified
   - `logTarget` parameter provides flexibility for where logs are sent

3. **Improved Error Handling** - Split the error handling into two separate sections for better diagnostics:
   - `RefreshFail` handler for errors during the actual refresh operation
   - `RefreshErrorHandler` for errors during setup or preparation

4. **Updated All Calling Code** - Modified `mod_510k_Processor.bas` to call the enhanced function with the correct connection base name "pgGet510kData".

## Code Implementation

The main changes in `mod_DataIO.RefreshPowerQuery`:

```vba
' Searching for the connection using multiple methods
For Each nameTry In Array(baseName, "Query - " & baseName, "Connection " & baseName)
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(nameTry)
    On Error GoTo 0
    If Not wbConn Is Nothing Then Exit For
Next nameTry

' If still not found, try loose search
If wbConn Is Nothing Then
    For Each c In ThisWorkbook.Connections
        If InStr(1, c.Name, baseName, vbTextCompare) > 0 Then
            Set wbConn = c
            Exit For
        End If
    Next c
End If
```

## Results

With this fix implemented:

1. The RunLog now shows "Found connection: Query - pgGet510kData" 
2. The WorkbookConnection refresh succeeds
3. Error #1004 no longer occurs
4. The scoring loop can proceed past AddScoreColumnsIfNeeded without errors

## Recommendations for Future Development

1. **Use Consistent Base Names** - Keep connection base names simple and predictable
2. **Reference by Pattern** - Remember that Excel automatically prepends "Query - " to Power Query connections
3. **Robust Connection Lookup** - Always implement robust connection lookup with multiple fallback methods
4. **Test Direct Connection Refresh** - When debugging, test direct connection refresh as shown above
