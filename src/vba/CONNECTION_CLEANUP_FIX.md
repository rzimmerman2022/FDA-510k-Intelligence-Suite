# Connection Cleanup Fix

## Issue Description

We were encountering a persistent Error #1004 during Power Query refresh operations. The root cause was identified as a sequence-of-operations issue in the refresh process:

1. In `RefreshPowerQuery`, we were calling `CleanupDuplicateConnections` first
2. This was removing ALL duplicate connections including the "good" one we needed
3. Later in the same function, when we tried to use the connection with `GetPQConnection`, it returned `Nothing`
4. This caused the code to fall back to `QueryTable.Refresh` which always fails with Error #1004 for Power Query tables

The log evidence showed this sequence clearly:
```
CleanupDuplicateConnections  INFO  Checking for duplicate connections based on found connection: 'Query - pgGet510kData'
RefreshPowerQuery            INFO  Attempting to get QueryTable from ListObject...
```

After cleanup, the connection was already gone, so it couldn't be found or used properly.

## Implemented Solution

The fix was two-fold:

1. **Updated `CleanupDuplicateConnections`** to accept parameters:
   - Added `connName` parameter to specify which base connection name to check (defaulting to "pgGet510kData")
   - Added `keepConn` parameter to specify a specific connection object that should be preserved

2. **Changed the sequence in `RefreshPowerQuery`**:
   - Get the connection object reference FIRST
   - Pass this connection to the cleanup function to ensure it's preserved
   - Perform the refresh through the preserved connection

## Implementation Details

### In CleanupDuplicateConnections:
- Added parameter validation to skip the connection specified by `keepConn`
- Improved connection counting to only remove true duplicates
- Enhanced logging to show how many duplicates were removed

### In RefreshPowerQuery:
- Restructured to get the workbook connection first
- Pass this connection to cleanup as the "keeper"
- Only after that, perform the refresh while we still have a valid connection

## Testing

This fix should resolve the Error #1004 issues during refresh. The expected sequence in logs should now show:

```
RefreshPowerQuery            INFO  Found WorkbookConnection: Query - pgGet510kData
CleanupDuplicateConnections  INFO  Using provided connection as keeper: 'Query - pgGet510kData'
RefreshPowerQuery            INFO  WorkbookConnection refreshed successfully
CleanupDuplicateConnections  INFO  Duplicate connections removed: X
```

## Additional Notes

The core insight was that we were cleaning up before getting a reference, essentially "shooting ourselves in the foot" by deleting the connection we needed. By reversing the operation sequence and preserving the active connection during cleanup, we maintain the integrity of the refresh process.
