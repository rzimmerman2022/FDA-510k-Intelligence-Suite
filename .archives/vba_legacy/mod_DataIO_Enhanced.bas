Attribute VB_Name = "mod_DataIO_Enhanced"
Option Explicit

' ==========================================================================
' Module      : mod_DataIO_Enhanced
' Author      : Cline (AI Assistant)
' Date        : May 7, 2025
' Purpose     : Enhanced implementation of data I/O functions, specifically
'               designed to address the Power Query refresh issues by using
'               timer-based execution contexts which have proven more reliable.
'
' Implementation Notes:
' This module provides a drop-in replacement for the RefreshPowerQuery function
' in mod_DataIO. The new implementation uses a timer-based approach that was 
' confirmed working during extensive testing. This approach breaks the execution 
' chain from the main workflow that was causing Error 1004.
'
' Dependencies: - mod_Logger: For logging I/O operations and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - Standard VBA libraries
'
' Tests Performed:
' 1. BackgroundQuery=False with Manual calculation - SUCCEEDED
' 2. BackgroundQuery=False with Automatic calculation - SUCCEEDED
' 3. BackgroundQuery=True with Manual calculation - SUCCEEDED
' 4. BackgroundQuery=True with Automatic calculation - FAILED (expected)
' 5. Direct VBA Call execution context - SUCCEEDED
' 6. CommandBar Button Click execution context - SUCCEEDED
' 7. Yes/No Dialog Response execution context - SUCCEEDED
' 8. Timer Event execution context - SUCCEEDED
'
' Based on these test results, we've determined that:
' - The Timer Event execution context is the most reliable
' - Setting BackgroundQuery=False is necessary
' - The issue is context-specific to the main workflow
' ==========================================================================

' --- Module-level variables ---
Private mTargetTableToRefresh As ListObject ' Holds the table being refreshed via timer
Private mRefreshTimer As Double             ' Holds the timer ID
Private mRefreshSuccess As Boolean          ' Holds the result of the async refresh operation
Private mRefreshAttemptCount As Integer     ' Counter for retries
Private mRefreshErrorMessage As String      ' Holds any error message
Private mRefreshWaitingForCompletion As Boolean ' Flag to indicate refresh is pending
Private Const MAX_REFRESH_ATTEMPTS As Integer = 2 ' Max number of retry attempts

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
    Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO_Enhanced.ExecuteTimerRefresh"
    
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
        Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO_Enhanced.ExecuteTimerRefresh"
    Else
        ' No more retries - update the module-level state
        LogEvt PROC_NAME, lgERROR, "All " & MAX_REFRESH_ATTEMPTS & " refresh attempts failed. Giving up."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "All refresh attempts failed", "MaxAttempts=" & MAX_REFRESH_ATTEMPTS
        mRefreshSuccess = False
        mRefreshWaitingForCompletion = False
    End If
End Sub

Public Sub CleanupDuplicateConnections()
    ' This calls the original CleanupDuplicateConnections in mod_DataIO
    mod_DataIO.CleanupDuplicateConnections
End Sub

' Other helpful utility functions

Public Function CancelRefresh() As Boolean
    ' Cancels any pending refresh operation
    On Error Resume Next
    
    If mRefreshTimer <> 0 Then
        Application.OnTime mRefreshTimer, "'" & ThisWorkbook.Name & "'!mod_DataIO_Enhanced.ExecuteTimerRefresh", , False
    End If
    
    mRefreshWaitingForCompletion = False
    mRefreshSuccess = False
    mRefreshErrorMessage = "Refresh cancelled by user"
    
    LogEvt "CancelRefresh", lgWARN, "Refresh operation cancelled by user"
    mod_DebugTraceHelpers.TraceEvt lvlWARN, "CancelRefresh", "Refresh cancelled by user"
    
    CancelRefresh = True
    On Error GoTo 0
End Function
