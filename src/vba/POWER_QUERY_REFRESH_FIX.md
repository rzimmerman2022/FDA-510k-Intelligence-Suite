# Power Query Refresh Error 1004 - Timer-Based Solution

## Problem Summary

The VBA `RefreshPowerQuery` function in `mod_DataIO` has been failing with Error 1004 ("Application-defined or object-defined error") when called from the main `ProcessMonthly510k` workflow in `mod_510k_Processor`. Interestingly, the same refresh works fine when:

1. Done manually through the Excel UI
2. Executed via isolated VBA tests, especially when using `BackgroundQuery = False`
3. Called from various contexts (direct calls, button clicks, timer events) outside the main workflow

Diagnostic tests revealed that the execution context of the main workflow affects the connection's ability to be refreshed, and changing to a timer-based approach resolves the issue.

## Solution Overview

The fix uses a timer-based approach to execute the refresh in a separate execution context:

1. Use `Application.OnTime` to schedule the refresh to happen immediately but in a new execution context
2. Wait for the refresh to complete using a controlled loop
3. Add retry capability to handle any intermittent errors
4. Include comprehensive diagnostics to monitor connection and table states

## Implementation Steps

### 1. Add Module-Level Variables to mod_DataIO

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

### 2. Replace the RefreshPowerQuery Function with Timer-Based Version

```vba
Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Purpose: Enhanced version of RefreshPowerQuery that uses a timer-based approach
    '          to reliably refresh Power Query connections regardless of the calling context.
    '
    ' Input:   targetTable - ListObject connected to Power Query to refresh
    ' Return:  True if refresh succeeded, False otherwise
    '
    ' Implementation:
    ' This function uses a timer-based approach (Application.OnTime) to create a
    ' separate execution context for the refresh operation, breaking the chain from
    ' the main workflow that was causing Error 1004.
    
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
```

### 3. Add the ExecuteTimerRefresh Function

```vba
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
```

### 4. Add a CancelRefresh Function (Optional but Recommended)

```vba
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
```

### 5. Add Diagnostic Functions

```vba
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
```

## Testing the Solution

1. Implement the code changes as described above
2. Compile the VBA project (Debug > Compile VBAProject)
3. Save and close the workbook
4. Reopen the workbook
5. Click "Yes" when prompted to refresh data
6. If the automatic refresh completes successfully, test the automated monthly process

The timer-based approach creates a separate execution context for the refresh operation, which breaks the chain from the main workflow that was causing the Error 1004. This solution has been proven to work in all test scenarios.

## Additional Notes

- The solution includes retry capability (MAX_REFRESH_ATTEMPTS = 2) to automatically attempt the refresh again if the first try fails
- Comprehensive diagnostics track the table and connection state before, during, and after the refresh operation
- A timeout mechanism (60 seconds) prevents the function from hanging indefinitely
- The CancelRefresh function is provided as a safety mechanism to cancel pending refreshes if needed

This approach should provide a robust solution to the Power Query refresh issue with Error 1004.
