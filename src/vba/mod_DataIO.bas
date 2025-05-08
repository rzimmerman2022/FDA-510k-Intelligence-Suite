Attribute VB_Name = "mod_DataIO"
Option Explicit

' ==========================================================================
' Module      : mod_DataIO
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : 2.0.0 (2025-05-07)
' ==========================================================================
' Description : This module handles data input/output operations for the
'               FDA 510(k) intelligence suite, including Power Query refresh,
'               table operations, and connection management.
'
' Key Feature : RefreshPowerQuery() - Enhanced with timer-based execution
'               context to solve the Error 1004 refresh issues. This version
'               uses Application.OnTime to create a separate execution context
'               that reliably refreshes Power Query connections regardless of
'               how it's called in the workflow.
'
' Dependencies: - mod_Logger: For logging I/O operations and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - Standard VBA libraries
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-05-07  Cline (AI)      - Completely revised RefreshPowerQuery to use
'                               a timer-based approach that resolves the 
'                               Error 1004 context-dependent issue
'                             - Added detailed diagnostics
'                             - Added automatic retry capability
'                             - Added comprehensive error logging
'                             - Enhanced connection cleanup
' [Previous dates/authors/changes unknown]
' ==========================================================================

' --- Constants ---
Private Const MODULE_NAME As String = "mod_DataIO"

' --- Module-level variables for timer-based refresh ---
Private mTargetTableToRefresh As ListObject ' Holds the table being refreshed via timer
Private mRefreshTimer As Double              ' Holds the timer ID
Private mRefreshSuccess As Boolean           ' Holds the result of the async refresh operation
Private mRefreshAttemptCount As Integer      ' Counter for retries
Private mRefreshErrorMessage As String       ' Holds any error message
Private mRefreshWaitingForCompletion As Boolean ' Flag to indicate refresh is pending
Private Const MAX_REFRESH_ATTEMPTS As Integer = 2 ' Max number of retry attempts

' ==========================================================================
' ===                    PRIMARY PUBLIC FUNCTIONS                        ===
' ==========================================================================

Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Purpose: Refreshes a Power Query table using a timer-based approach that
    '          creates a separate execution context, resolving Error 1004 issues.
    '
    ' Input:   targetTable - ListObject connected to Power Query to refresh
    ' Return:  True if refresh succeeded, False otherwise
    '
    ' Implementation:
    ' This function uses Application.OnTime to create a separate execution context
    ' for the refresh operation, breaking the chain from the main workflow
    ' that was causing Error 1004.
    
    Const PROC_NAME As String = "RefreshPowerQuery"
    RefreshPowerQuery = False ' Default to failure
    
    ' Validate input
    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot refresh."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If
    
    ' Detailed diagnostics on the input table
    DiagnoseTableState targetTable, "Input-Check", PROC_NAME
    
    ' Store the target table in module-level variable for the timer function to use
    Set mTargetTableToRefresh = targetTable
    
    ' Reset module-level variables
    mRefreshSuccess = False
    mRefreshAttemptCount = 0
    mRefreshErrorMessage = ""
    mRefreshWaitingForCompletion = True
    
    ' Log that we're initiating the timer-based refresh
    LogEvt PROC_NAME, lgINFO, "Initiating timer-based refresh for table: " & targetTable.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Initiating timer-based refresh", "Table=" & targetTable.Name
    
    ' Schedule the refresh to happen via timer (immediately but in new execution context)
    mRefreshTimer = Now ' Current time
    Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO.ExecuteTimerRefresh"
    
    ' Wait for the refresh to complete
    Dim waitStart As Double
    waitStart = Timer
    
    ' Display a status message while waiting
    Application.StatusBar = "Refreshing Power Query data... Please wait."
    
    ' Wait for completion or timeout after 60 seconds
    Do While mRefreshWaitingForCompletion
        DoEvents ' Allow Excel to process events, including our timer
        
        ' Check for timeout (60 seconds)
        If Timer - waitStart > 60 Then
            Application.StatusBar = False
            LogEvt PROC_NAME, lgERROR, "Refresh operation timed out after 60 seconds"
            mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Refresh operation timed out", "Table=" & targetTable.Name
            mRefreshWaitingForCompletion = False
            mRefreshSuccess = False
            mRefreshErrorMessage = "Refresh operation timed out after 60 seconds"
        End If
        
        ' Brief pause to prevent CPU spinning
        Application.Wait Now + TimeSerial(0, 0, 0.1) ' 0.1 second pause
    Loop
    
    ' Clear the status bar
    Application.StatusBar = False
    
    ' Return the result
    RefreshPowerQuery = mRefreshSuccess
    
    ' If the refresh failed, log and display the error message
    If Not mRefreshSuccess Then
        LogEvt PROC_NAME, lgERROR, "Timer-based refresh failed. " & mRefreshErrorMessage
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Timer-based refresh failed", "Error=" & mRefreshErrorMessage
    Else
        LogEvt PROC_NAME, lgINFO, "Timer-based refresh completed successfully for table: " & targetTable.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Timer-based refresh succeeded", "Table=" & targetTable.Name
        
        ' Clean up any potential duplicate connections
        Call CleanupDuplicateConnections
    End If
    
    ' Clean up
    Set mTargetTableToRefresh = Nothing
    
    Exit Function
End Function

Public Sub ExecuteTimerRefresh()
    ' Purpose: This is the function called by the Application.OnTime timer
    '          to perform the actual refresh in a separate execution context.
    '
    ' Note: This should not be called directly - it's triggered via timer
    
    Const PROC_NAME As String = "ExecuteTimerRefresh"
    Dim wbConn As WorkbookConnection
    Dim qt As QueryTable
    Dim connName As String
    Dim connNameAlt As String
    Dim success As Boolean
    Dim originalCalc As XlCalculation
    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    
    ' Initialize
    success = False
    connName = "pgGet510kData"
    connNameAlt = "Query - " & connName
    
    ' Ensure module variables are valid
    If mTargetTableToRefresh Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing in timer callback"
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table is Nothing in timer callback"
        mRefreshSuccess = False
        mRefreshErrorMessage = "Internal error: Target table was lost"
        mRefreshWaitingForCompletion = False
        Exit Sub
    End If
    
    ' Increment attempt counter
    mRefreshAttemptCount = mRefreshAttemptCount + 1
    
    LogEvt PROC_NAME, lgINFO, "Timer-based refresh attempt " & mRefreshAttemptCount & " for table: " & mTargetTableToRefresh.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Timer refresh attempt " & mRefreshAttemptCount, "Table=" & mTargetTableToRefresh.Name
    
    ' Store original environment settings
    originalCalc = Application.Calculation
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    
    ' Create clean environment for refresh
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Check target table state again in this context
    DiagnoseTableState mTargetTableToRefresh, "Timer-Context", PROC_NAME
    
    ' First try to find the associated QueryTable
    Set qt = Nothing
    On Error Resume Next
    Set qt = mTargetTableToRefresh.QueryTable
    
    ' Find the connection - first try the direct table's connection
    Set wbConn = Nothing
    On Error Resume Next
    
    ' Try finding the connection by common names
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt)
    
    ' If still not found and we have a QueryTable, try using its connection
    If wbConn Is Nothing And Not qt Is Nothing Then
        Dim connString As String
        connString = qt.Connection
        
        ' Extract connection name from connection string if possible
        ' This is a simplistic approach - connection strings vary
        If InStr(connString, "Connection=") > 0 Then
            Dim startPos As Integer, endPos As Integer
            startPos = InStr(connString, "Connection=") + Len("Connection=")
            endPos = InStr(startPos, connString, ";")
            If endPos > startPos Then
                connName = Mid(connString, startPos, endPos - startPos)
                Set wbConn = ThisWorkbook.Connections(connName)
            End If
        End If
    End If
    
    ' If connection still not found, also try harder by checking all connections
    If wbConn Is Nothing Then
        Dim c As WorkbookConnection
        For Each c In ThisWorkbook.Connections
            ' Look for connections containing the base name
            If InStr(c.Name, "pgGet510kData") > 0 Then
                Set wbConn = c
                Exit For
            End If
        Next c
    End If
    
    On Error GoTo RefreshErrorHandler
    
    ' Verify we found a connection
    If wbConn Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Could not find WorkbookConnection for refresh"
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Could not find WorkbookConnection"
        Err.Raise Number:=9999, Description:="Could not find WorkbookConnection for refresh"
    End If
    
    ' Log found connection
    LogEvt PROC_NAME, lgINFO, "Found connection: " & wbConn.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found connection", "Name=" & wbConn.Name
    
    ' Detailed connection state diagnostics
    DiagnoseConnectionState wbConn, "Pre-Refresh", PROC_NAME
    
    ' Ensure the workbook is active
    ThisWorkbook.Activate
    
    ' Clear message queue
    DoEvents
    DoEvents
    
    ' Set connection properties for reliable refresh
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
        On Error Resume Next
        wbConn.OLEDBConnection.EnableRefresh = True ' Ensure refresh is enabled
        On Error GoTo RefreshErrorHandler
    End If
    
    ' Perform the refresh
    LogEvt PROC_NAME, lgINFO, "Executing refresh on connection: " & wbConn.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Executing refresh", "Connection=" & wbConn.Name
    
    wbConn.Refresh
    
    ' If no error occurred, refresh was successful
    success = True
    LogEvt PROC_NAME, lgINFO, "Refresh command succeeded for connection: " & wbConn.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Refresh command succeeded", "Connection=" & wbConn.Name
    
    ' One more diagnostic check after successful refresh
    DiagnoseConnectionState wbConn, "Post-Refresh", PROC_NAME
    
    ' Clean up and exit
    Set wbConn = Nothing
    Set qt = Nothing
    
    ' Restore environment
    Application.Calculation = originalCalc
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Update module-level state
    mRefreshSuccess = success
    mRefreshWaitingForCompletion = False
    Exit Sub
    
RefreshErrorHandler:
    ' Handle any errors
    success = False
    mRefreshErrorMessage = "Error " & Err.Number & ": " & Err.Description
    
    LogEvt PROC_NAME, lgERROR, "Error during refresh: " & mRefreshErrorMessage, "Connection=" & IIf(wbConn Is Nothing, "Unknown", wbConn.Name)
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during refresh", "Error=" & mRefreshErrorMessage
    
    ' If we have a connection, check its state after error
    If Not wbConn Is Nothing Then
        DiagnoseConnectionState wbConn, "Post-Error", PROC_NAME
    End If
    
    ' Clean up
    Set wbConn = Nothing
    Set qt = Nothing
    
    ' Restore environment
    Application.Calculation = originalCalc
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Should we retry?
    If mRefreshAttemptCount < MAX_REFRESH_ATTEMPTS Then
        LogEvt PROC_NAME, lgWARN, "Refresh attempt " & mRefreshAttemptCount & " failed. Will retry..."
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Will retry refresh", "Attempt=" & mRefreshAttemptCount
        
        ' Schedule a retry in 2 seconds
        mRefreshTimer = Now + TimeSerial(0, 0, 2)
        Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO.ExecuteTimerRefresh"
    Else
        ' No more retries - update the module-level state
        LogEvt PROC_NAME, lgERROR, "All " & MAX_REFRESH_ATTEMPTS & " refresh attempts failed. Giving up."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "All refresh attempts failed", "MaxAttempts=" & MAX_REFRESH_ATTEMPTS
        mRefreshSuccess = False
        mRefreshWaitingForCompletion = False
    End If
End Sub

Public Function CancelRefresh() As Boolean
    ' Cancels any pending refresh operation
    On Error Resume Next
    
    If mRefreshTimer <> 0 Then
        Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO.ExecuteTimerRefresh", , False
    End If
    
    mRefreshWaitingForCompletion = False
    mRefreshSuccess = False
    mRefreshErrorMessage = "Refresh cancelled by user"
    
    LogEvt "CancelRefresh", lgWARN, "Refresh operation cancelled by user"
    mod_DebugTraceHelpers.TraceEvt lvlWARN, "CancelRefresh", "Refresh cancelled by user"
    
    CancelRefresh = True
    On Error GoTo 0
End Function

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

Public Function SheetExists(sheetName As String) As Boolean
    ' Purpose: Checks if a worksheet with the given name exists in the workbook.
    ' Returns: True if the sheet exists, False otherwise.
    Dim ws As Worksheet
    
    SheetExists = False ' Default to not found
    
    On Error Resume Next ' Ignore errors if sheet not found
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0 ' Restore default error handling
End Function

' ==========================================================================
' ===                    CONNECTION MANAGEMENT                           ===
' ==========================================================================

Public Sub CleanupDuplicateConnections()
    ' Purpose: Safely cleans up duplicate or orphaned connections
    ' This version uses a two-phase approach to avoid collection modification issues:
    ' 1. First identify all connections to delete
    ' 2. Then delete them (in reverse order) after identification is complete
    
    Const PROC_NAME As String = "CleanupDuplicateConnections"
    
    ' Variables for connection processing
    Dim c As WorkbookConnection
    Dim connectionsToDelete As Collection
    Dim i As Integer, j As Integer
    Dim connectionList As String
    Dim connCount As Integer
    
    ' Initialize collection to store connections for deletion
    Set connectionsToDelete = New Collection
    
    ' Log starting connection count
    connCount = ThisWorkbook.Connections.Count
    LogEvt PROC_NAME, lgINFO, "Starting connection cleanup. Initial count: " & connCount
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Starting cleanup", "Count=" & connCount
    
    ' Skip processing if no connections exist
    If connCount = 0 Then
        LogEvt PROC_NAME, lgINFO, "No connections found - nothing to clean up."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "No connections found"
        Exit Sub
    End If
    
    ' ===== PHASE 1: Identify connections to delete =====
    On Error Resume Next
    For i = 1 To ThisWorkbook.Connections.Count
        Set c = ThisWorkbook.Connections(i)
        
        If Not c Is Nothing Then
            ' Log each connection for debugging
            LogEvt PROC_NAME, lgDETAIL, "Connection " & i & ": " & c.Name
            
            ' Check for duplicate or stale connections based on naming patterns
            ' and other criteria
            
            ' Example: Check for connections with duplicate prefix/naming patterns
            If InStr(1, c.Name, "Query - pgGet510kData") > 0 Then
                ' This is likely a duplicate PQ connection
                LogEvt PROC_NAME, lgDETAIL, "Marking for deletion: " & c.Name & " (duplicate prefix)"
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Marked for deletion", "Name=" & c.Name & ", Reason=duplicate prefix"
                connectionsToDelete.Add i, CStr(i)
            ElseIf InStr(1, c.Name, "_xlnm.") > 0 Then
                ' This is an Excel-generated connection that might be stale
                LogEvt PROC_NAME, lgDETAIL, "Marking for deletion: " & c.Name & " (Excel-generated)"
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Marked for deletion", "Name=" & c.Name & ", Reason=Excel-generated"
                connectionsToDelete.Add i, CStr(i)
            End If
            
            ' Add additional criteria as needed based on your specific requirements
            ' For example, checking connection state, target, etc.
        End If
    Next i
    On Error GoTo 0
    
    ' ===== PHASE 2: Delete identified connections (in reverse order) =====
    If connectionsToDelete.Count > 0 Then
        LogEvt PROC_NAME, lgINFO, "Found " & connectionsToDelete.Count & " connections to remove."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Connections to remove", "Count=" & connectionsToDelete.Count
        
        ' Process in REVERSE order to avoid index shifting problems
        For i = connectionsToDelete.Count To 1 Step -1
            On Error Resume Next
            j = connectionsToDelete(i) ' Get the stored index
            
            ' Get connection name before deletion (for logging)
            Dim connName As String
            connName = "Unknown"
            If j <= ThisWorkbook.Connections.Count Then
                If Not ThisWorkbook.Connections(j) Is Nothing Then
                    connName = ThisWorkbook.Connections(j).Name
                End If
            End If
            
            ' Delete the connection
            ThisWorkbook.Connections(j).Delete
            
            ' Log result
            If Err.Number = 0 Then
                LogEvt PROC_NAME, lgINFO, "Successfully removed connection: " & connName
                mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Removed connection", "Name=" & connName
            Else
                LogEvt PROC_NAME, lgWARN, "Failed to remove connection " & connName & ". Error: " & Err.Number & " - " & Err.Description
                mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Failed to remove connection", "Name=" & connName & ", Err=" & Err.Number
            End If
            Err.Clear
            On Error GoTo 0
        Next i
    Else
        LogEvt PROC_NAME, lgINFO, "No connections identified for removal."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "No connections to remove"
    End If
    
    ' Log final connection count
    LogEvt PROC_NAME, lgINFO, "Cleanup complete. Final count: " & ThisWorkbook.Connections.Count
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Cleanup complete", "FinalCount=" & ThisWorkbook.Connections.Count
    
    ' Clean up
    Set connectionsToDelete = Nothing
End Sub

' ==========================================================================
' ===                    DIAGNOSTIC FUNCTIONS                            ===
' ==========================================================================

Private Sub DiagnoseTableState(tbl As ListObject, phase As String, proc As String)
    ' Purpose: Comprehensive diagnostics for a ListObject's state
    ' This function checks and logs detailed information about a table object
    
    On Error Resume Next
    
    ' Basic table properties
    LogEvt proc, lgDETAIL, phase & " - Table Diagnostic: Name=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Table Check", "Name=" & tbl.Name
    
    ' Parent worksheet info
    If tbl.Parent Is Nothing Then
        LogEvt proc, lgWARN, phase & " - Table parent worksheet is Nothing!"
        mod_DebugTraceHelpers.TraceEvt lvlWARN, proc, phase & " - Table Parent Missing", "Table=" & tbl.Name
    Else
        LogEvt proc, lgDETAIL, phase & " - Table parent worksheet: " & tbl.Parent.Name
        mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Table Parent", "Sheet=" & tbl.Parent.Name
    End If
    
    ' Table range info
    On Error Resume Next
    Dim rangeAddr As String
    rangeAddr = "Error getting range"
    If Not tbl.Range Is Nothing Then
        rangeAddr = tbl.Range.Address(External:=True)
    End If
    LogEvt proc, lgDETAIL, phase & " - Table range: " & rangeAddr
    mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Table Range", "Address=" & rangeAddr
    
    ' QueryTable information
    Dim tempQT As QueryTable
    Set tempQT = Nothing
    
    On Error Resume Next
    Set tempQT = tbl.QueryTable
    
    ' Check QueryTable existence
    If Err.Number <> 0 Then
        LogEvt proc, lgWARN, phase & " - Error checking QueryTable: " & Err.Number & " - " & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlWARN, proc, phase & " - QueryTable Error", "Err=" & Err.Number & " - " & Err.Description
        Err.Clear
    ElseIf tempQT Is Nothing Then
        LogEvt proc, lgWARN, phase & " - Table.QueryTable is Nothing!"
        mod_DebugTraceHelpers.TraceEvt lvlWARN, proc, phase & " - QueryTable Missing"
    Else
        ' QueryTable exists and is valid
        Dim qtProps As String
        qtProps = "Name=" & tempQT.Name
        
        ' Check refreshing state
        On Error Resume Next
        qtProps = qtProps & ", Refreshing=" & tempQT.Refreshing
        If Err.Number <> 0 Then qtProps = qtProps & " (Error checking Refreshing property)"
        Err.Clear
        
        ' Check background query setting
        On Error Resume Next
        qtProps = qtProps & ", BackgroundQuery=" & tempQT.BackgroundQuery
        If Err.Number <> 0 Then qtProps = qtProps & " (Error checking BackgroundQuery property)"
        Err.Clear
        
        ' Log QueryTable info
        LogEvt proc, lgDETAIL, phase & " - QueryTable properties: " & qtProps
        mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - QueryTable Properties", qtProps
    End If
    
    ' Application state
    Dim appState As String
    appState = "Calculation=" & Application.Calculation & ", ScreenUpdating=" & Application.ScreenUpdating & ", EnableEvents=" & Application.EnableEvents
    LogEvt proc, lgDETAIL, phase & " - Application state: " & appState
    mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Application State", appState
    
    ' Clean up
    Set tempQT = Nothing
    On Error GoTo 0
End Sub

Private Sub DiagnoseConnectionState(conn As WorkbookConnection, phase As String, proc As String)
    ' Purpose: Comprehensive diagnostics for a WorkbookConnection's state
    ' This function checks and logs detailed information about a connection object
    
    On Error Resume Next
    
    ' Basic connection properties
    Dim connDetails As String
    connDetails = "Name=" & conn.Name & ", Type="
    
    ' Connection type
    Select Case conn.Type
        Case xlConnectionTypeOLEDB: connDetails = connDetails & "OLEDB"
        Case xlConnectionTypeODBC: connDetails = connDetails & "ODBC"
        Case xlConnectionTypeXMLMAP: connDetails = connDetails & "XMLMAP"
        Case xlConnectionTypeTEXT: connDetails = connDetails & "TEXT"
        Case xlConnectionTypeWEB: connDetails = connDetails & "WEB"
        Case Else: connDetails = connDetails & "OTHER-" & conn.Type
    End Select
    
    LogEvt proc, lgDETAIL, phase & " - Connection: " & connDetails
    mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Connection Basic Info", connDetails
    
    ' Check refreshing state
    On Error Resume Next
    LogEvt proc, lgDETAIL, phase & " - Connection.Refreshing=" & conn.Refreshing
    mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - Connection Refreshing State", "Refreshing=" & conn.Refreshing
    If Err.Number <> 0 Then
        LogEvt proc, lgWARN, phase & " - Error checking Refreshing: " & Err.Number & " - " & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlWARN, proc, phase & " - Error checking Refreshing", "Err=" & Err.Number
    End If
    Err.Clear
    
    ' For OLEDB connections, get additional details
    If conn.Type = xlConnectionTypeOLEDB Then
        Dim oledbInfo As String
        oledbInfo = ""
        
        ' Check MaintainConnection
        On Error Resume Next
        oledbInfo = oledbInfo & "MaintainConnection=" & conn.OLEDBConnection.MaintainConnection
        If Err.Number <> 0 Then
            oledbInfo = oledbInfo & " (Error " & Err.Number & " checking MaintainConnection)"
            Err.Clear
        End If
        
        ' Check BackgroundQuery
        On Error Resume Next
        oledbInfo = oledbInfo & ", BackgroundQuery=" & conn.OLEDBConnection.BackgroundQuery
        If Err.Number <> 0 Then
            oledbInfo = oledbInfo & " (Error " & Err.Number & " checking BackgroundQuery)"
            Err.Clear
        End If
        
        ' Check EnableRefresh
        On Error Resume Next
        oledbInfo = oledbInfo & ", EnableRefresh=" & conn.OLEDBConnection.EnableRefresh
        If Err.Number <> 0 Then
            oledbInfo = oledbInfo & " (Error " & Err.Number & " checking EnableRefresh)"
            Err.Clear
        End If
        
        ' Check IsConnected
        On Error Resume Next
        oledbInfo = oledbInfo & ", IsConnected="
        Dim isConnected As Boolean
        isConnected = conn.OLEDBConnection.IsConnected
        oledbInfo = oledbInfo & isConnected
        If Err.Number <> 0 Then
            oledbInfo = oledbInfo & " (Error " & Err.Number & " checking IsConnected)"
            Err.Clear
        End If
        
        ' Log OLEDB-specific info
        LogEvt proc, lgDETAIL, phase & " - OLEDBConnection: " & oledbInfo
        mod_DebugTraceHelpers.TraceEvt lvlDET, proc, phase & " - OLEDB Details", oledbInfo
    End If
    
    On Error GoTo 0
End Sub
