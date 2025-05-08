# Power Query Refresh After Archive Fix

## Issue Summary

When trying to manually refresh the Power Query via VBA after archiving, users encountered the following error:
```
Error #1004: Application-defined or object-defined error
```

The error occurred because:
1. The archive process was converting the `pgGet510kData` table to a static range using `ListObject.Unlist`
2. This disconnected the table from its `QueryTable` object
3. When a subsequent refresh attempt was made, `RefreshPowerQuery` failed because `QueryTable` was `Nothing`

## Solution Implemented

The fix preserves the original table while still creating an archive copy by:
1. **Creating a values-only copy** of the data rather than unlisting the original table
2. **Preserving the original QueryTable connection** which allows subsequent PQ refreshes to succeed

### Files Modified

1. **mod_Archive.bas**
   - Modified `ArchiveIfNeeded` to create a new sheet with values-only data instead of converting the original table
   - Removed call to `mod_DataIO.CleanupDuplicateConnections` as it's no longer needed
   - Updated documentation to reflect the changes

2. **mod_DataIO.bas**
   - Updated documentation for `CleanupDuplicateConnections` to note it's no longer called during archiving
   - The function remains available for other purposes or manual cleanup

3. **mod_510k_Processor.bas**
   - Added a comment about preserving the original table when calling `ArchiveIfNeeded`

### Before and After Behavior

**Before:**
1. Archive process copied the entire sheet
2. Called `tblArchive.Unlist` to convert table to static range
3. Called `CleanupDuplicateConnections` to handle duplicated connections
4. Subsequent PQ refresh failed with Error #1004 because QueryTable was disconnected

**After:**
1. Archive process creates a **new sheet**
2. Copies **data as values-only** to this new sheet
3. **Preserves original table structure** and QueryTable connection
4. Subsequent PQ refresh succeeds because QueryTable is still connected

## Testing

To test the fix:
1. Run a full processing cycle including archiving (`ProcessMonthly510k`)
2. After completion, run the process again
3. Verify that the Power Query refresh succeeds with no Error #1004

## Benefits

1. **More Reliable Processing**: Eliminates cryptic Error #1004 when running the process multiple times
2. **Cleaner Architecture**: Archive is now a true "snapshot" that doesn't affect the active data table
3. **Reduced Complexity**: No need to clean up duplicate connections, as we're not copying sheets with queries

## Implementation Date
May 8, 2025
