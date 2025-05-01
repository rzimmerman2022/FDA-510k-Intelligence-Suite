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
'               - RefreshPowerQuery: Refreshes the specified ListObject's
'                 associated QueryTable synchronously and attempts to disable
'                 background refresh afterward.
'               - SheetExists: Checks if a sheet with a given name exists.
'               - CleanupDuplicateConnections: Attempts to identify and delete
'                 duplicate Power Query connections based on naming patterns,
'                 often needed after sheet copying.
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
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Ensure EnableRefresh is True before attempting refresh
'                               in RefreshPowerQuery to fix "Refresh disabled" error.
' 2025-04-30  Cline (AI)      - Made EnableRefresh check more explicit in RefreshPowerQuery
'                               and removed On Error Resume Next around setting it True.
' 2025-04-30  Cline (AI)      - Added simple TestCall function for debugging compile error.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit
Attribute VB_Name = "mod_DataIO"

Public Function TestCall() As Boolean
    ' Simple test function
    TestCall = True
End Function

Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Purpose: Refreshes the Power Query associated with the target table using QueryTable object.
    '          Includes disabling background refresh post-query.
    Dim qt As QueryTable
    Const PROC_NAME As String = "RefreshPowerQuery"
    RefreshPowerQuery = False ' Default to False

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot refresh."
        TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    On Error GoTo RefreshErrorHandler
    LogEvt PROC_NAME, lgINFO, "Attempting QueryTable refresh for: " & targetTable.Name
    TraceEvt lvlINFO, PROC_NAME, "Start refresh", "Table='" & targetTable.Name & "'"

    Set qt = targetTable.QueryTable
    If qt Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Could not find QueryTable associated with table '" & targetTable.Name & "'."
        TraceEvt lvlERROR, PROC_NAME, "QueryTable object is Nothing", "Table='" & targetTable.Name & "'"
        MsgBox "Error: Could not find QueryTable associated with table '" & targetTable.Name & "'.", vbCritical, "Refresh Error"
        Exit Function ' Exit, cannot refresh
    End If

    ' --- Ensure refresh is enabled BEFORE attempting ---
    If Not qt.EnableRefresh Then
        LogEvt PROC_NAME, lgWARN, "QueryTable refresh was disabled for '" & targetTable.Name & "'. Attempting to enable.", "Current EnableRefresh=" & qt.EnableRefresh
        TraceEvt lvlWARN, PROC_NAME, "Refresh was disabled. Attempting enable.", "Table='" & targetTable.Name & "'"
        ' Removed On Error Resume Next here to catch potential errors setting the property
        qt.EnableRefresh = True
        ' Re-check if it succeeded
        If Not qt.EnableRefresh Then
             LogEvt PROC_NAME, lgERROR, "Failed to set EnableRefresh=True for table '" & targetTable.Name & "'. Halting refresh attempt."
             TraceEvt lvlERROR, PROC_NAME, "Failed to set EnableRefresh=True", "Table='" & targetTable.Name & "'"
             MsgBox "Error: Could not enable refresh for table '" & targetTable.Name & "'.", vbCritical, "Refresh Error"
             Exit Function ' Cannot proceed
        Else
             LogEvt PROC_NAME, lgINFO, "Successfully set EnableRefresh=True for table '" & targetTable.Name & "'."
             TraceEvt lvlINFO, PROC_NAME, "Successfully set EnableRefresh=True", "Table='" & targetTable.Name & "'"
        End If
    Else
        LogEvt PROC_NAME, lgDETAIL, "QueryTable refresh already enabled for '" & targetTable.Name & "'.", "Current EnableRefresh=" & qt.EnableRefresh
        TraceEvt lvlDET, PROC_NAME, "Refresh already enabled", "Table='" & targetTable.Name & "'"
    End If

    ' Refresh synchronously
    qt.BackgroundQuery = False ' Ensure background query is off for synchronous refresh
    qt.Refresh

    ' --- Lock refresh settings post-query (per review suggestion) ---
    On Error Resume Next ' Best effort to disable these
    qt.BackgroundQuery = False
    qt.EnableRefresh = False
    If Err.Number <> 0 Then
         LogEvt PROC_NAME, lgWARN, "Could not disable BackgroundQuery/EnableRefresh after refresh for table '" & targetTable.Name & "'. Error: " & Err.Description
         TraceEvt lvlWARN, PROC_NAME, "Failed to set BackgroundQuery=False / EnableRefresh=False", "Table='" & targetTable.Name & "', Err=" & Err.Description
         Err.Clear
    Else
         LogEvt PROC_NAME, lgDETAIL, "Set BackgroundQuery=False and EnableRefresh=False post-refresh for table '" & targetTable.Name & "'."
         TraceEvt lvlDET, PROC_NAME, "Set BackgroundQuery=False / EnableRefresh=False post-refresh", "Table='" & targetTable.Name & "'"
    End If
    On Error GoTo RefreshErrorHandler ' Restore main handler for this sub
    ' --- End Lock ---

    RefreshPowerQuery = True ' If refresh completes without error
    LogEvt PROC_NAME, lgINFO, "QueryTable refresh completed successfully for: " & targetTable.Name
    TraceEvt lvlINFO, PROC_NAME, "Refresh successful", "Table='" & targetTable.Name & "'"
    Exit Function

RefreshErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    RefreshPowerQuery = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error during QueryTable refresh for '" & targetTable.Name & "'. Error #" & errNum & ": " & errDesc
    TraceEvt lvlERROR, PROC_NAME, "Error during QueryTable refresh", "Table='" & targetTable.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "QueryTable refresh failed for table '" & targetTable.Name & "': " & vbCrLf & errDesc, vbExclamation, "Refresh Error"
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
    ' Purpose: Removes duplicate Power Query connections often created by sheet copying (e.g., during archiving).
    Const PROC_NAME As String = "CleanupDuplicateConnections"
    Dim c As WorkbookConnection
    Dim baseConnectionName As String
    Dim originalConnection As WorkbookConnection ' To store ref if found
    Set originalConnection = Nothing

    TraceEvt lvlINFO, PROC_NAME, "Phase: Cleanup Connections Start"

    ' Try to find the original connection by common names
    On Error Resume Next
    Set originalConnection = ThisWorkbook.Connections("pgGet510kData")
    If originalConnection Is Nothing Then Set originalConnection = ThisWorkbook.Connections("Query - pgGet510kData")
    On Error GoTo 0 ' Restore default error handling

    If Not originalConnection Is Nothing Then
        baseConnectionName = originalConnection.Name
        LogEvt PROC_NAME, lgINFO, "Checking for duplicate connections based on found connection: '" & baseConnectionName & "'"
        TraceEvt lvlINFO, PROC_NAME, "Checking duplicate connections", "Base=" & baseConnectionName
        On Error Resume Next ' Ignore errors during loop/delete
        For Each c In ThisWorkbook.Connections
            ' Check if name is different AND follows the "Base Name (Number)" pattern
            If c.Name <> baseConnectionName And c.Name Like baseConnectionName & " (*" Then
                LogEvt PROC_NAME, lgDETAIL, "Deleting duplicate connection: " & c.Name
                TraceEvt lvlDET, PROC_NAME, "Deleting duplicate connection", c.Name
                c.Delete
            End If
        Next c
        On Error GoTo 0 ' Restore default error handling
    Else
         ' Fallback if original connection wasn't found by name
         baseConnectionName = "pgGet510kData" ' Assume this is the most likely base
         LogEvt PROC_NAME, lgWARN, "Could not find original PQ connection by typical names. Attempting cleanup based on pattern: '" & baseConnectionName & " (*' or 'Query - " & baseConnectionName & " (*'"
         TraceEvt lvlWARN, PROC_NAME, "Original PQ Connection not found", "FallbackBase=" & baseConnectionName
         On Error Resume Next ' Ignore errors during loop/delete
         For Each c In ThisWorkbook.Connections
             If c.Name Like baseConnectionName & " (*" Or c.Name Like "Query - " & baseConnectionName & " (*" Then
                 LogEvt PROC_NAME, lgDETAIL, "Deleting potential duplicate connection: " & c.Name
                 TraceEvt lvlDET, PROC_NAME, "Deleting potential duplicate connection", c.Name
                 c.Delete
             End If
         Next c
         On Error GoTo 0 ' Restore default error handling
    End If

    Set c = Nothing
    Set originalConnection = Nothing
    TraceEvt lvlINFO, PROC_NAME, "Phase: Cleanup Connections End"
End Sub

Public Function ArrayToTable(dataArr As Variant, targetTable As ListObject) As Boolean
    ' Purpose: Writes a 2D data array back to the DataBodyRange of a target ListObject.
    ' Returns: True if successful, False otherwise.
    Const PROC_NAME As String = "ArrayToTable"
    ArrayToTable = False ' Default to failure

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot write array."
        TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    If Not IsArray(dataArr) Then
        LogEvt PROC_NAME, lgERROR, "Input dataArr is not a valid array. Cannot write."
        TraceEvt lvlERROR, PROC_NAME, "Input dataArr is not an array", "Table=" & targetTable.Name
        Exit Function
    End If

    On Error GoTo WriteErrorHandler

    Dim numRows As Long, numCols As Long
    On Error Resume Next ' Check array bounds safely
    numRows = UBound(dataArr, 1) - LBound(dataArr, 1) + 1
    numCols = UBound(dataArr, 2) - LBound(dataArr, 2) + 1
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Error getting bounds of dataArr. Cannot write.", "Table=" & targetTable.Name & ", Err=" & Err.Description
        TraceEvt lvlERROR, PROC_NAME, "Error getting array bounds", "Table=" & targetTable.Name & ", Err=" & Err.Description
        Err.Clear
        On Error GoTo WriteErrorHandler ' Restore handler
        Exit Function
    End If
    On Error GoTo WriteErrorHandler ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Attempting to write array (" & numRows & "x" & numCols & ") to table '" & targetTable.Name & "'."
    TraceEvt lvlDET, PROC_NAME, "Start writing array to table", "Table=" & targetTable.Name & ", Size=" & numRows & "x" & numCols

    ' --- Resize table if necessary (optional, but safer) ---
    ' Clear existing data first to avoid issues if new array is smaller
    If targetTable.ListRows.Count > 0 Then
        targetTable.DataBodyRange.ClearContents
    End If
    ' Resize based on array dimensions (if table allows resizing)
    ' Note: Resizing might fail if table is linked externally in certain ways.
    ' Consider adding more robust resizing logic if needed.
    ' For now, assume direct write is sufficient if dimensions match or table auto-expands.

    ' --- Write the array ---
    targetTable.DataBodyRange.Resize(numRows, numCols).Value = dataArr

    ArrayToTable = True ' Success
    LogEvt PROC_NAME, lgINFO, "Successfully wrote array to table '" & targetTable.Name & "'."
    TraceEvt lvlINFO, PROC_NAME, "Array write successful", "Table=" & targetTable.Name
    Exit Function

WriteErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    ArrayToTable = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error writing array to table '" & targetTable.Name & "'. Error #" & errNum & ": " & errDesc
    TraceEvt lvlERROR, PROC_NAME, "Error writing array to table", "Table='" & targetTable.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "Error writing data back to table '" & targetTable.Name & "': " & vbCrLf & errDesc, vbExclamation, "Write Error"
    ' Exit Function ' Exit implicitly after error handler
End Function
