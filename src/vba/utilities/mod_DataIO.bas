' ==========================================================================
' Module      : mod_DataIO
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module encapsulates data input/output operations for the
'               application. Its primary responsibilities include refreshing
'               Power Query data sources associated with specified tables,
'               checking if worksheets exist, writing VBA array data back into
'               Excel ListObjects (tables), and cleaning up potentially
'               duplicated Power Query connections that can occur when worksheets
'               containing queries are copied (e.g., during archiving).
'
' Key Functions:
'               - RefreshPowerQuery: Refreshes the target table's associated
'                 QueryTable synchronously. Duplicate columns handled later.
'               - ClearExistingTable: Clears range and deletes a ListObject (Currently Unused).
'               - SheetExists: Checks if a sheet with a given name exists.
'               - CleanupDuplicateConnections: Attempts to identify and delete
'                 duplicate Power Query connections based on naming patterns.
'                 (Note: Previously used after sheet copying during archiving,
'                 but no longer needed with the new archive implementation.)
'               - ArrayToTable: Writes the contents of a 2D VBA array into
'                 the DataBodyRange of a specified ListObject.
'
' Dependencies: - mod_Logger: For logging I/O operations and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-05-11  Cline (AI)      - Fixed critical PowerQuery refresh issue by reversing operations
'                               order: grab connection reference first, THEN clean up duplicates
'                               while preserving the connection reference, THEN refresh.
'                             - Updated CleanupDuplicateConnections to accept keepConn parameter
'                               to prevent deleting the active connection during cleanup.
'                             - Fixed bug where we deleted connection then tried to use it.
'                             - Added connection counting and improved duplicate logic.
'                             - Created CONNECTION_CLEANUP_FIX.md documentation.
' 2025-05-08  Cline (AI)      - Fixed Error #1004 during refresh by changing RefreshPowerQuery
'                               to use WorkbookConnection.Refresh as primary path,
'                               with QueryTable.Refresh as fallback.
'                             - Added automatic call to CleanupDuplicateConnections before
'                               each refresh to prevent duplicate query errors.
'                             - Removed setting EnableRefresh = False after refresh
'                               as this can invalidate the connection.
'                             - Updated CleanupDuplicateConnections documentation to
'                               note it is no longer needed by the archive implementation
'                               due to changes in mod_Archive (values-only copy).
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Ensure EnableRefresh is True before attempting refresh
'                               in RefreshPowerQuery to fix "Refresh disabled" error.
' 2025-04-30  Cline (AI)      - Made EnableRefresh check more explicit in RefreshPowerQuery
'                               and removed On Error Resume Next around setting it True.
' 2025-04-30  Cline (AI)      - Added simple TestCall function for debugging compile error.
' 2025-04-30  Cline (AI)      - Added ClearExistingTable sub and call it in RefreshPowerQuery
'                               to prevent duplicate columns from header collisions.
' 2025-04-30  Cline (AI)      - Modified ClearExistingTable to clear contents/headers only,
'                               NOT delete the ListObject, to fix "Object required" error.
'                             - Simplified RefreshPowerQuery to use the existing QueryTable
'                               object associated with the non-deleted ListObject.
' 2025-04-30  Cline (AI)      - Renamed ClearExistingTable to ClearExistingTableRows.
'                             - Modified ClearExistingTableRows to delete DataBodyRange only.
'                             - Updated RefreshPowerQuery to call ClearExistingTableRows.
' 2025-04-30  Cline (AI)      - Changed ClearExistingTableRows to use .ClearContents
'                               instead of .Delete to fix error 1004 on refresh.
' 2025-04-30  Cline (AI)      - Reverted ClearExistingTableRows back to ClearExistingTable
'                               (including .Delete) and modified RefreshPowerQuery to
'                               refresh via WorkbookConnection object instead of QueryTable.
' 2025-04-30  Cline (AI)      - Removed call to ClearExistingTable from RefreshPowerQuery
'                               to avoid pre-refresh table manipulation errors. Duplicate
'                               columns will be handled post-refresh by mod_Format.
' 2025-04-30  Cline (AI)      - Qualified all TraceEvt calls with mod_DebugTraceHelpers.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit


Public Function TestCall() As Boolean
    ' Simple test function
    TestCall = True
End Function

Public Sub ClearExistingTable(lo As ListObject)
    ' Purpose: Clears all data, formatting, and headers from a ListObject's range
    '          and then deletes the ListObject itself. Prepares for connection refresh.
    '          NOTE: This is currently unused as pre-refresh deletion caused errors.
    Const PROC_NAME As String = "ClearExistingTable"
    If lo Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Target table object is Nothing. Cannot clear."
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Target table object is Nothing"
        Exit Sub
    End If

    On Error GoTo ClearErrorHandler
    LogEvt PROC_NAME, lgINFO, "Attempting to clear range and delete table: " & lo.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start clear range/delete", "Table=" & lo.Name

    ' Use With block for clarity
    With lo
        ' Clear contents, formats, headers etc. from the entire range
        .Range.Clear
        ' Delete the ListObject itself
        .Delete ' Reinstated delete
    End With

    LogEvt PROC_NAME, lgINFO, "Successfully cleared range and deleted table."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Clear range/delete successful"
    Exit Sub

ClearErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    LogEvt PROC_NAME, lgERROR, "Error clearing range/deleting table '" & lo.Name & "'. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error clearing range/deleting table", "Table='" & lo.Name & "', Err=" & errNum & " - " & errDesc
    ' Optionally re-raise or handle differently, but for now just log
End Sub


Public Function RefreshPowerQuery(targetTable As ListObject, Optional baseName As String = "", _
                                  Optional logTarget As Object = Nothing) As Boolean
    ' Purpose: Refreshes the Power Query associated with the target table. 
    '          First attempts via WorkbookConnection (modern approach), 
    '          falls back to QueryTable if necessary (legacy approach).
    '          Duplicate columns may occur and should be handled post-refresh by mod_Format.
    Dim qt As QueryTable
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim nameTry As Variant, c As WorkbookConnection
    Const PROC_NAME As String = "RefreshPowerQuery"
    RefreshPowerQuery = False ' Default to False

    ' Use LogTarget or fall back to default logging
    If logTarget Is Nothing Then
        ' Use normal LogEvt and TraceEvt
        LogEvt PROC_NAME, lgINFO, "Starting refresh with default logging"
    End If

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot refresh."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If
    
    ' If baseName not provided, use target table name
    If baseName = "" Then
        baseName = targetTable.Name
    End If

    On Error GoTo RefreshErrorHandler

    '==============================
    ' 1. Try likely exact matches
    '==============================
    LogEvt PROC_NAME, lgINFO, "Looking for WorkbookConnection with base name: " & baseName
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Looking for connection", "BaseName=" & baseName
    
    For Each nameTry In Array(baseName, "Query - " & baseName, "Connection " & baseName)
        On Error Resume Next
        Set wbConn = ThisWorkbook.Connections(nameTry)
        On Error GoTo 0
        If Not wbConn Is Nothing Then 
            LogEvt PROC_NAME, lgINFO, "Found connection: " & wbConn.Name
            mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found connection by exact name", "Name=" & wbConn.Name
            Exit For
        End If
    Next nameTry
    
    '==============================
    ' 2. Fallback: loose search
    '==============================
    If wbConn Is Nothing Then
        LogEvt PROC_NAME, lgINFO, "No exact match found. Trying loose search for: " & baseName
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Trying loose search", "Pattern=" & baseName
        
        For Each c In ThisWorkbook.Connections
            If InStr(1, c.Name, baseName, vbTextCompare) > 0 Then
                Set wbConn = c
                LogEvt PROC_NAME, lgINFO, "Found connection via loose search: " & wbConn.Name
                mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found connection by pattern match", "Name=" & wbConn.Name
                Exit For
            End If
        Next c
    End If
    
    ' Fall back to QueryTable only if no WorkbookConnection found
    If wbConn Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "No WorkbookConnection found. Falling back to QueryTable."
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "No connection found, trying QueryTable fallback"
        
        Set qt = targetTable.QueryTable
        If qt Is Nothing Then
            LogEvt PROC_NAME, lgERROR, "Could not find QueryTable associated with table '" & targetTable.Name & "'."
            mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "QueryTable object is Nothing", "Table='" & targetTable.Name & "'"
            MsgBox "Error: Could not find the Power Query connection associated with table '" & targetTable.Name & "'. Cannot refresh.", vbCritical, "Refresh Error"
            Exit Function ' Exit, cannot refresh
        End If
        
        connName = qt.Name
        LogEvt PROC_NAME, lgINFO, "Found QueryTable: " & connName
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found QueryTable", "Name=" & connName
    
        ' --- Ensure refresh is enabled BEFORE attempting ---
        If Not qt.EnableRefresh Then
            LogEvt PROC_NAME, lgWARN, "QueryTable refresh was disabled for '" & connName & "'. Attempting to enable."
            mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Refresh was disabled. Attempting enable.", "QueryTable='" & connName & "'"
            qt.EnableRefresh = True
            If Not qt.EnableRefresh Then
                 LogEvt PROC_NAME, lgERROR, "Failed to set EnableRefresh=True for QueryTable '" & connName & "'. Halting refresh attempt."
                 mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Failed to set EnableRefresh=True", "QueryTable='" & connName & "'"
                 MsgBox "Error: Could not enable refresh for QueryTable '" & connName & "'.", vbCritical, "Refresh Error"
                 Exit Function ' Cannot proceed
            Else
                 LogEvt PROC_NAME, lgINFO, "Successfully set EnableRefresh=True for QueryTable '" & connName & "'."
                 mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Successfully set EnableRefresh=True", "QueryTable='" & connName & "'"
            End If
        End If
    End If

    '==============================
    ' 3. Refresh via appropriate method
    '==============================
    On Error GoTo RefreshFail
    
    If Not wbConn Is Nothing Then
        ' --- MODERN PATH: Refresh via WorkbookConnection (preferred) ---
        LogEvt PROC_NAME, lgINFO, "Refreshing via WorkbookConnection: " & wbConn.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Starting WorkbookConnection refresh", "Connection=" & wbConn.Name
        
        ' --- Set background query to false for synchronous refresh ---
        If Not wbConn.OLEDBConnection Is Nothing Then
            wbConn.OLEDBConnection.BackgroundQuery = False
        End If
        
        ' --- Clean up any duplicate connections BEFORE refreshing ---
        If baseName <> "" Then
            CleanupDuplicateConnections baseName, wbConn
            LogEvt PROC_NAME, lgINFO, "Cleaned up any duplicate connections except our primary connection"
            mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Ran CleanupDuplicateConnections", "KeptConnection=" & wbConn.Name
        End If
        
        ' --- Perform the actual refresh ---
        wbConn.Refresh ' <<< REFRESH VIA WORKBOOK CONNECTION (MODERN PATH) >>>
        
        LogEvt PROC_NAME, lgINFO, "Refresh succeeded via WorkbookConnection"
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Refresh successful via WorkbookConnection", "Connection=" & wbConn.Name
    Else
        ' --- LEGACY PATH: Fall back to QueryTable refresh ---
        LogEvt PROC_NAME, lgINFO, "Attempting synchronous refresh via QueryTable: " & connName
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Starting QueryTable refresh", "QueryTable=" & connName
        
        qt.BackgroundQuery = False ' Ensure synchronous refresh
        qt.Refresh BackgroundQuery:=False ' <<< REFRESH VIA QUERYTABLE (LEGACY PATH) >>>
        
        LogEvt PROC_NAME, lgINFO, "Refresh succeeded via QueryTable"
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Refresh successful via QueryTable", "QueryTable=" & connName
    End If

    RefreshPowerQuery = True ' If we got here, refresh completed successfully
    Exit Function

RefreshFail:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    LogEvt PROC_NAME, lgERROR, "Error during refresh. Err #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during refresh", "Err=" & errNum & " - " & errDesc
    RefreshPowerQuery = False
    Exit Function
    
RefreshErrorHandler:
    errDesc = Err.Description
    errNum = Err.Number
    RefreshPowerQuery = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error during refresh process. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during refresh process", "Err=" & errNum & " - " & errDesc
    MsgBox "Refresh failed: " & vbCrLf & errDesc, vbExclamation, "Refresh Error"
End Function

Public Function SheetExists(sheetName As String) As Boolean
    ' Purpose: Checks if a sheet with the given name exists in the workbook.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
    Set ws = Nothing ' Clean up
End Function

Public Sub CleanupDuplicateConnections(Optional connName As String = "pgGet510kData", _
                                      Optional keepConn As WorkbookConnection = Nothing)
    ' Purpose: Removes duplicate Power Query connections that may be created by various operations.
    ' Note: This function was previously called by archiving process, but is no longer needed
    '       after the May 2025 update to mod_Archive that uses values-only copies.
    ' Parameters:
    '   connName - Base name of the connection to check for duplicates
    '   keepConn - Optional specific connection object to preserve (won't be deleted)
    ' The function remains available for manual cleanup or other code that may need it.
    Const PROC_NAME As String = "CleanupDuplicateConnections"
    Dim c As WorkbookConnection
    Dim baseConnectionName As String
    Dim originalConnection As WorkbookConnection ' To store ref if found
    Dim n As Long ' Counter for connections with the base name
    Set originalConnection = Nothing

    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Phase: Cleanup Connections Start"

    ' If we were passed a specific connection to keep, use that
    If Not keepConn Is Nothing Then
        baseConnectionName = connName
        Set originalConnection = keepConn
        LogEvt PROC_NAME, lgINFO, "Using provided connection as keeper: '" & keepConn.Name & "'"
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Using provided keeper", "Keeper=" & keepConn.Name
    Else
        ' Try to find the original connection by common names
        On Error Resume Next
        Set originalConnection = ThisWorkbook.Connections(connName)
        If originalConnection Is Nothing Then Set originalConnection = ThisWorkbook.Connections("Query - " & connName)
        On Error GoTo 0 ' Restore default error handling
    End If

    If Not originalConnection Is Nothing Then
        baseConnectionName = connName
        LogEvt PROC_NAME, lgINFO, "Checking for duplicate connections based on found connection: '" & baseConnectionName & "'"
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Checking duplicate connections", "Base=" & baseConnectionName
        On Error Resume Next ' Ignore errors during loop/delete
        For Each c In ThisWorkbook.Connections
            ' Skip the connection we're keeping
            If Not keepConn Is Nothing Then
                If c.Name = keepConn.Name Then
                    LogEvt PROC_NAME, lgDETAIL, "Keeping connection: " & c.Name
                    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Keeping connection", "Name=" & c.Name
                    GoTo NextConn
                End If
            End If
            
            ' Check if name contains the base name (using pattern matching)
            If InStr(1, c.Name, baseConnectionName, vbTextCompare) > 0 Then
                n = n + 1
                If n > 1 Then ' Only delete duplicates, not the first one found
                    LogEvt PROC_NAME, lgDETAIL, "Deleting duplicate connection: " & c.Name
                    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting duplicate connection", c.Name
                    c.Delete
                End If
            End If
NextConn:
        Next c
        On Error GoTo 0 ' Restore default error handling
    Else
         ' Fallback if original connection wasn't found by name
         baseConnectionName = connName ' Use the provided name or default
         LogEvt PROC_NAME, lgWARN, "Could not find original PQ connection by typical names. Attempting cleanup based on pattern: '" & baseConnectionName & "'"
         mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Original PQ Connection not found", "FallbackBase=" & baseConnectionName
         On Error Resume Next ' Ignore errors during loop/delete
         For Each c In ThisWorkbook.Connections
             ' Skip the keeper connection if specified
             If Not keepConn Is Nothing Then
                 If c.Name = keepConn.Name Then GoTo NextConnFallback
             End If
            
             If InStr(1, c.Name, baseConnectionName, vbTextCompare) > 0 Then
                 n = n + 1
                 If n > 1 Then ' Only delete duplicates, not the first one found
                     LogEvt PROC_NAME, lgDETAIL, "Deleting potential duplicate connection: " & c.Name
                     mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting potential duplicate connection", c.Name
                     c.Delete
                 End If
             End If
NextConnFallback:
         Next c
         On Error GoTo 0 ' Restore default error handling
    End If

    LogEvt PROC_NAME, lgINFO, "Duplicate connections removed: " & (n - 1)
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Connections cleanup complete", "DuplicatesRemoved=" & (n - 1)
    
    Set c = Nothing
    Set originalConnection = Nothing
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Phase: Cleanup Connections End"
End Sub

Public Function ArrayToTable(dataArr As Variant, targetTable As ListObject) As Boolean
    ' Purpose: Writes a 2D data array back to the DataBodyRange of a target ListObject.
    ' Returns: True if successful, False otherwise.
    Const PROC_NAME As String = "ArrayToTable"
    ArrayToTable = False ' Default to failure

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot write array."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    If Not IsArray(dataArr) Then
        LogEvt PROC_NAME, lgERROR, "Input dataArr is not a valid array. Cannot write."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Input dataArr is not an array", "Table=" & targetTable.Name
        Exit Function
    End If

    On Error GoTo WriteErrorHandler

    Dim numRows As Long, numCols As Long
    On Error Resume Next ' Check array bounds safely
    numRows = UBound(dataArr, 1) - LBound(dataArr, 1) + 1
    numCols = UBound(dataArr, 2) - LBound(dataArr, 2) + 1
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Error getting bounds of dataArr. Cannot write.", "Table=" & targetTable.Name & ", Err=" & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error getting array bounds", "Table=" & targetTable.Name & ", Err=" & Err.Description
        Err.Clear
        On Error GoTo WriteErrorHandler ' Restore handler
        Exit Function
    End If
    On Error GoTo WriteErrorHandler ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Attempting to write array (" & numRows & "x" & numCols & ") to table '" & targetTable.Name & "'."
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start writing array to table", "Table=" & targetTable.Name & ", Size=" & numRows & "x" & numCols

    ' --- Resize table if necessary (optional, but safer) ---
    ' Clear existing data first to avoid issues if new array is smaller
    If targetTable.ListRows.Count > 0 Then
        targetTable.DataBodyRange.ClearContents ' Changed from .Delete to .ClearContents
    End If
    ' Resize based on array dimensions (if table allows resizing)
    ' Note: Resizing might fail if table is linked externally in certain ways.
    ' Consider adding more robust resizing logic if needed.
    ' For now, assume direct write is sufficient if dimensions match or table auto-expands.

    ' --- Write the array ---
    targetTable.DataBodyRange.Resize(numRows, numCols).value = dataArr

    ArrayToTable = True ' Success
    LogEvt PROC_NAME, lgINFO, "Successfully wrote array to table '" & targetTable.Name & "'."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Array write successful", "Table=" & targetTable.Name
    Exit Function

WriteErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    ArrayToTable = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error writing array to table '" & targetTable.Name & "'. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error writing array to table", "Table='" & targetTable.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "Error writing data back to table '" & targetTable.Name & "': " & vbCrLf & errDesc, vbExclamation, "Write Error"
    ' Exit Function ' Exit implicitly after error handler
End Function
