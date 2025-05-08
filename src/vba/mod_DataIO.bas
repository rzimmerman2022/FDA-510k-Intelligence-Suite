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
' 2025-05-08  Cline (AI)      - Updated CleanupDuplicateConnections documentation to
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


Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Purpose: Refreshes the Power Query associated with the target table using QueryTable object.
    '          Duplicate columns may occur and should be handled post-refresh by mod_Format.
    Dim qt As QueryTable
    Const PROC_NAME As String = "RefreshPowerQuery"
    RefreshPowerQuery = False ' Default to False

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot refresh."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    On Error GoTo RefreshErrorHandler

    ' --- Get the QueryTable directly from the ListObject ---
    LogEvt PROC_NAME, lgINFO, "Attempting to get QueryTable from ListObject: " & targetTable.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Getting QueryTable from ListObject", "Table=" & targetTable.Name
    Set qt = targetTable.QueryTable

    If qt Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Could not find QueryTable associated with table '" & targetTable.Name & "'."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "QueryTable object is Nothing", "Table='" & targetTable.Name & "'"
        MsgBox "Error: Could not find the Power Query connection associated with table '" & targetTable.Name & "'. Cannot refresh.", vbCritical, "Refresh Error"
        Exit Function ' Exit, cannot refresh
    End If
    LogEvt PROC_NAME, lgINFO, "Found QueryTable: " & qt.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found QueryTable", "Name=" & qt.Name

    ' --- Ensure refresh is enabled BEFORE attempting ---
    If Not qt.EnableRefresh Then
        LogEvt PROC_NAME, lgWARN, "QueryTable refresh was disabled for '" & qt.Name & "'. Attempting to enable.", "Current EnableRefresh=" & qt.EnableRefresh
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Refresh was disabled. Attempting enable.", "QueryTable='" & qt.Name & "'"
        qt.EnableRefresh = True
        If Not qt.EnableRefresh Then
             LogEvt PROC_NAME, lgERROR, "Failed to set EnableRefresh=True for QueryTable '" & qt.Name & "'. Halting refresh attempt."
             mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Failed to set EnableRefresh=True", "QueryTable='" & qt.Name & "'"
             MsgBox "Error: Could not enable refresh for QueryTable '" & qt.Name & "'.", vbCritical, "Refresh Error"
             Exit Function ' Cannot proceed
        Else
             LogEvt PROC_NAME, lgINFO, "Successfully set EnableRefresh=True for QueryTable '" & qt.Name & "'."
             mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Successfully set EnableRefresh=True", "QueryTable='" & qt.Name & "'"
        End If
    Else
        LogEvt PROC_NAME, lgDETAIL, "QueryTable refresh already enabled for '" & qt.Name & "'.", "Current EnableRefresh=" & qt.EnableRefresh
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Refresh already enabled", "QueryTable='" & qt.Name & "'"
    End If

    ' --- Find the Workbook Connection ---
    Dim wbConn As WorkbookConnection
    Dim connName As String: connName = qt.Name ' Start assuming connection name matches QueryTable name
    Dim connNameAlt As String: connNameAlt = "Query - " & qt.Name ' Alternative name format

    On Error Resume Next ' Try finding connection by primary or alternative name
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt)
    On Error GoTo RefreshErrorHandler ' Restore main handler

    If wbConn Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Could not find WorkbookConnection named '" & connName & "' or '" & connNameAlt & "'."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "WorkbookConnection not found", "TriedName1=" & connName & ", TriedName2=" & connNameAlt
        MsgBox "Error: Could not find the Workbook Connection associated with QueryTable '" & qt.Name & "'. Cannot refresh.", vbCritical, "Refresh Error"
        Exit Function
    End If
    LogEvt PROC_NAME, lgINFO, "Found WorkbookConnection: " & wbConn.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found WorkbookConnection", "Name=" & wbConn.Name

    ' --- Refresh via QueryTable ---
    ' Reverted from wbConn.Refresh back to qt.Refresh 2025-04-30 per user feedback on refresh error
    qt.BackgroundQuery = False ' Ensure synchronous refresh
    LogEvt PROC_NAME, lgINFO, "Attempting synchronous refresh via QueryTable: " & qt.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Starting synchronous refresh via QueryTable", "QueryTable=" & qt.Name
    qt.Refresh BackgroundQuery:=False ' <<< REFRESH VIA QUERYTABLE >>>

    ' --- Lock refresh settings post-query ---
    ' Re-enable QueryTable locking as it controls sheet interaction
    On Error Resume Next ' Best effort to disable these
    qt.BackgroundQuery = False ' Should already be false, but ensure
    qt.EnableRefresh = False
     If Err.Number <> 0 Then
         LogEvt PROC_NAME, lgWARN, "Could not disable BackgroundQuery/EnableRefresh on QueryTable '" & qt.Name & "' after refresh. Error: " & Err.Description
         mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Failed to set QueryTable BackgroundQuery=False / EnableRefresh=False post-refresh", "QueryTable='" & qt.Name & "', Err=" & Err.Description
         Err.Clear
    Else
         LogEvt PROC_NAME, lgDETAIL, "Set QueryTable BackgroundQuery=False and EnableRefresh=False post-refresh for '" & qt.Name & "'."
         mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Set QueryTable BackgroundQuery=False / EnableRefresh=False post-refresh", "QueryTable='" & qt.Name & "'"
    End If
    On Error GoTo RefreshErrorHandler ' Restore main handler

    RefreshPowerQuery = True ' If refresh completes without error
    LogEvt PROC_NAME, lgINFO, "QueryTable refresh completed successfully for: " & qt.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Refresh successful via QueryTable", "QueryTable=" & qt.Name
    Exit Function

RefreshErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    RefreshPowerQuery = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error during QueryTable refresh process. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during QueryTable refresh process", "Err=" & errNum & " - " & errDesc
    MsgBox "QueryTable refresh failed: " & vbCrLf & errDesc, vbExclamation, "Refresh Error" ' Updated MsgBox text
    ' Exit Function ' Exit implicitly after error handler
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

Public Sub CleanupDuplicateConnections()
    ' Purpose: Removes duplicate Power Query connections that may be created by various operations.
    ' Note: This function was previously called by archiving process, but is no longer needed
    '       after the May 2025 update to mod_Archive that uses values-only copies.
    ' The function remains available for manual cleanup or other code that may need it.
    Const PROC_NAME As String = "CleanupDuplicateConnections"
    Dim c As WorkbookConnection
    Dim baseConnectionName As String
    Dim originalConnection As WorkbookConnection ' To store ref if found
    Set originalConnection = Nothing

    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Phase: Cleanup Connections Start"

    ' Try to find the original connection by common names
    On Error Resume Next
    Set originalConnection = ThisWorkbook.Connections("pgGet510kData")
    If originalConnection Is Nothing Then Set originalConnection = ThisWorkbook.Connections("Query - pgGet510kData")
    On Error GoTo 0 ' Restore default error handling

    If Not originalConnection Is Nothing Then
        baseConnectionName = originalConnection.Name
        LogEvt PROC_NAME, lgINFO, "Checking for duplicate connections based on found connection: '" & baseConnectionName & "'"
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Checking duplicate connections", "Base=" & baseConnectionName
        On Error Resume Next ' Ignore errors during loop/delete
        For Each c In ThisWorkbook.Connections
            ' Check if name is different AND follows the "Base Name (Number)" pattern
            If c.Name <> baseConnectionName And c.Name Like baseConnectionName & " (*" Then
                LogEvt PROC_NAME, lgDETAIL, "Deleting duplicate connection: " & c.Name
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting duplicate connection", c.Name
                c.Delete
            End If
        Next c
        On Error GoTo 0 ' Restore default error handling
    Else
         ' Fallback if original connection wasn't found by name
         baseConnectionName = "pgGet510kData" ' Assume this is the most likely base
         LogEvt PROC_NAME, lgWARN, "Could not find original PQ connection by typical names. Attempting cleanup based on pattern: '" & baseConnectionName & " (*' or 'Query - " & baseConnectionName & " (*'"
         mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Original PQ Connection not found", "FallbackBase=" & baseConnectionName
         On Error Resume Next ' Ignore errors during loop/delete
         For Each c In ThisWorkbook.Connections
             If c.Name Like baseConnectionName & " (*" Or c.Name Like "Query - " & baseConnectionName & " (*" Then
                 LogEvt PROC_NAME, lgDETAIL, "Deleting potential duplicate connection: " & c.Name
                 mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting potential duplicate connection", c.Name
                 c.Delete
             End If
         Next c
         On Error GoTo 0 ' Restore default error handling
    End If

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
