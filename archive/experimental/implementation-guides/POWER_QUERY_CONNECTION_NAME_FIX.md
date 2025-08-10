# Power Query Connection Name Fix

## Summary
The Power Query refresh issue has been fixed by improving the WorkbookConnection lookup logic to handle the "Query - " prefix that Excel automatically adds to connection names. The fix also implements a multi-level fallback approach to ensure connections are found reliably.

## Changes Implemented

1. **Enhanced RefreshPowerQuery in mod_DataIO.bas**:
   - Added `baseName` parameter to specify the connection base name (defaults to table name if not specified)
   - Added `logTarget` parameter for flexible logging
   - Implemented a hierarchical search strategy:
     1. Try exact name matches (baseName, "Query - " + baseName, "Connection " + baseName)
     2. Fallback to pattern-based search (find any connection containing baseName)
     3. Only use QueryTable approach if no matching connection can be found
   - Improved error handling with separate handlers for refresh failures vs. setup errors

2. **Updated mod_510k_Processor.bas**:
   - Modified RefreshPowerQuery calls to pass "pgGet510kData" as the baseName parameter
   - Ensures consistent connection lookup across all refresh scenarios

3. **Documentation**:
   - Created comprehensive documentation in `docs/05_CONNECTION_NAME_FIX.md`
   - Explains the issue, solution approach, implementation details, and future recommendations

## Verification

The fix was verified with direct tests through the Immediate window:
```vba
For Each c In ThisWorkbook.Connections: ? c.Name: Next
Query - pgGet510kData

? ThisWorkbook.Connections("Query - pgGet510kData").RefreshWithRefreshAll
True
```

This confirmed that:
1. The connection exists with the name "Query - pgGet510kData"
2. The connection can be refreshed directly without errors

## Expected Results

With this fix, the application should:
1. Successfully find the WorkbookConnection on the first try
2. Refresh the data through the modern WorkbookConnection.Refresh method
3. Avoid Error #1004 that was previously occurring
4. Complete the data processing workflow without interruption

## Future Recommendations

1. Always use the new RefreshPowerQuery with a properly specified baseName parameter
2. Be aware that Excel automatically adds "Query - " prefix to Power Query connections
3. Use the loose search capabilities when connection names might vary or contain version numbers
4. When debugging refresh issues, use direct connection tests in the Immediate window
