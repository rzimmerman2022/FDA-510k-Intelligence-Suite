Option Explicit

' ==========================================================================
' Module      : mod_RefreshSolutions
' Author      : Cline (AI Assistant)
' Date        : May 7, 2025
' Purpose     : Provides potential alternative implementation solutions for
'               the Power Query refresh operation based on different
'               execution context patterns to address the issue where refresh
'               fails when triggered automatically but works manually.
'
' HOW TO USE  : Once you've run the diagnostic tests in mod_TestRefresh and
'               mod_TestWithContext, you can determine which of these solutions
'               to try by observing which execution contexts succeeded and
'               which failed. Each solution addresses different refresh
'               failure patterns and root causes.
' ==========================================================================

' Solution 1: Async Refresh via OnTime Timer
' Best for: When the YES/NO Dialog context fails but direct calls succeed
' Description: This separates the MsgBox action from the refresh operation
'              by scheduling the refresh to occur in a new execution context
'              using Application.OnTime.
Public Sub Solution1_AsyncRefreshViaTimer()
    Dim connName As String
    
    connName = "pgGet510kData" ' Update if your connection has a different name
    
    ' When used in a message box context like:
    ' If MsgBox("Refresh data?", vbYesNo) = vbYes Then
    '     Call this function instead of directly refreshing
    ' End If
    
    ' Schedule the refresh to happen immediately but in a new execution context
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!mod_RefreshSolutions.PerformDelayedRefresh"
End Sub

Public Sub PerformDelayedRefresh()
    ' This gets executed in a separate context from the MsgBox response
    Dim wbConn As WorkbookConnection
    
    On Error Resume Next
    
    ' Try to find the connection
    Set wbConn = ThisWorkbook.Connections("pgGet510kData")
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections("Query - pgGet510kData")
    
    If wbConn Is Nothing Then
        MsgBox "Could not find Power Query connection to refresh.", vbExclamation, "Refresh Error"
        Exit Sub
    End If
    
    ' Prepare the environment
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure foreground refresh
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    
    ' Important: Make sure the workbook is active
    ThisWorkbook.Activate
    
    ' Clear the message queue
    DoEvents
    
    ' Perform the refresh
    wbConn.Refresh
    
    ' Check for error
    If Err.Number <> 0 Then
        MsgBox "Error refreshing data: " & Err.Description, vbCritical, "Refresh Error"
    End If
    
    ' Restore environment
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set wbConn = Nothing
End Sub

' Solution 2: Message Queue Clearing
' Best for: When refresh issues appear to be related to the Excel message queue
' Description: This uses multiple DoEvents calls to ensure the message queue
'              is completely processed before and after attempting refresh.
Public Function Solution2_MessageQueueClearing(targetTable As ListObject) As Boolean
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim connNameAlt As String
    Dim success As Boolean
    
    ' Initialize
    Solution2_MessageQueueClearing = False
    connName = "pgGet510kData"
    connNameAlt = "Query - " & connName
    
    ' Ensure message queue is processed
    DoEvents
    DoEvents ' Double DoEvents can help clear deeply queued messages
    
    ' Find the connection
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt)
    
    If wbConn Is Nothing Then
        MsgBox "Could not find Power Query connection to refresh.", vbExclamation, "Refresh Error"
        Exit Function
    End If
    
    ' Prepare the environment
    Dim originalCalc As XlCalculation
    originalCalc = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure foreground refresh
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    
    ' Clear message queue again
    DoEvents
    
    ' Refresh
    wbConn.Refresh
    success = (Err.Number = 0)
    
    ' Cleanup
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Process any pending messages
    DoEvents
    
    Solution2_MessageQueueClearing = success
    
    Set wbConn = Nothing
End Function

' Solution 3: Force Environment Reset
' Best for: When refresh fails in multiple contexts and may be related to Excel state
' Description: This approach forcibly resets multiple Excel environment settings
'              and performs extra steps to ensure a clean state before refresh.
Public Function Solution3_ForceEnvironmentReset(targetTable As ListObject) As Boolean
    Dim wbConn As WorkbookConnection
    Dim qt As QueryTable
    Dim connName As String
    Dim connNameAlt As String
    Dim originalCalc As XlCalculation
    Dim success As Boolean
    
    ' Initialize
    Solution3_ForceEnvironmentReset = False
    connName = "pgGet510kData"
    connNameAlt = "Query - " & connName
    
    ' Save current state
    originalCalc = Application.Calculation
    
    ' Full environment reset
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Wait briefly to ensure environment settings are applied
    Application.Wait Now + TimeSerial(0, 0, 1) ' 1 second delay
    
    ' Ensure the workbook is active
    ThisWorkbook.Activate
    DoEvents
    
    ' Find the connection multiple ways
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt)
    
    ' If still nothing, try to find via the query table
    If wbConn Is Nothing And Not targetTable Is Nothing Then
        Set qt = targetTable.QueryTable
        If Not qt Is Nothing Then
            ' Try to get the connection from the query table
            connName = qt.Connection
            Set wbConn = ThisWorkbook.Connections(connName)
        End If
    End If
    
    If wbConn Is Nothing Then
        MsgBox "Could not find Power Query connection to refresh.", vbExclamation, "Refresh Error"
        GoTo CleanupAndExit
    End If
    
    ' Force connection properties
    If Not wbConn.OLEDBConnection Is Nothing Then
        ' Reset connection properties
        wbConn.OLEDBConnection.BackgroundQuery = False
        
        ' Attempt to ensure refresh is enabled
        On Error Resume Next
        wbConn.OLEDBConnection.EnableRefresh = True
        If Err.Number <> 0 Then
            Err.Clear
            ' Try alternate approach if EnableRefresh fails
        End If
    End If
    
    ' Clear Excel's message queue thoroughly
    DoEvents
    DoEvents
    
    ' Perform the refresh with additional error handling
    On Error Resume Next
    wbConn.Refresh
    
    ' Check for success
    If Err.Number <> 0 Then
        Debug.Print "Refresh error: " & Err.Description
        success = False
    Else
        success = True
    End If
    
CleanupAndExit:
    ' Restore environment (ordered for most stability)
    Application.Calculation = originalCalc
    DoEvents
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Clean up objects
    If Not qt Is Nothing Then Set qt = Nothing
    Set wbConn = Nothing
    
    Solution3_ForceEnvironmentReset = success
End Function

' Solution 4: Hybrid Approach - UI Response with Timer Delay
' Best for: Complex UI interactions where timing and context are both issues
' Description: This solution combines UI prompting with a timer delay and
'              careful environment preparation.
Public Function Solution4_PromptWithTimerRefresh(promptMessage As String) As Boolean
    ' Initialize
    Solution4_PromptWithTimerRefresh = False
    
    ' Process any pending events before showing prompt
    DoEvents
    
    ' Show the prompt
    If MsgBox(promptMessage, vbQuestion + vbYesNo, "Refresh Data") = vbYes Then
        ' User clicked Yes - schedule a slightly delayed refresh
        ' This breaks the direct execution chain from MsgBox to refresh
        Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!mod_RefreshSolutions.ExecuteRefreshWithRetry"
        Solution4_PromptWithTimerRefresh = True ' Return True since user said Yes
    End If
End Function

Public Sub ExecuteRefreshWithRetry()
    ' This is called via timer to perform the refresh with retry capability
    Dim wbConn As WorkbookConnection
    Dim retry As Boolean
    Dim attemptCount As Integer
    
    ' Initialize
    retry = False
    attemptCount = 0
    
    ' Reset Excel environment
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Activate
    DoEvents
    
    ' Find the connection
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections("pgGet510kData")
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections("Query - pgGet510kData")
    
    If wbConn Is Nothing Then
        MsgBox "Could not find Power Query connection to refresh.", vbExclamation, "Refresh Error"
        GoTo Cleanup
    End If
    
    ' Ensure foreground refresh
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    
RetryRefresh:
    ' Attempt the refresh
    On Error Resume Next
    wbConn.Refresh
    
    ' Check if it worked
    If Err.Number <> 0 Then
        attemptCount = attemptCount + 1
        Debug.Print "Refresh attempt " & attemptCount & " failed: " & Err.Description
        
        ' Only retry once
        If attemptCount = 1 Then
            ' Clear error and try again with a small delay
            Err.Clear
            Application.Wait Now + TimeSerial(0, 0, 2) ' 2 second delay
            retry = True
            GoTo RetryRefresh
        Else
            ' Give up after second attempt
            MsgBox "Error refreshing data: " & Err.Description, vbCritical, "Refresh Error"
        End If
    Else
        Debug.Print "Refresh succeeded on attempt " & attemptCount + 1
    End If
    
Cleanup:
    ' Restore environment
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set wbConn = Nothing
End Sub

' Solution 5: WorkbookConnection State Reset
' Best for: When connection state may be corrupted
' Description: This approach attempts a connection cleanup and state reset
Public Function Solution5_ConnectionStateReset(targetTable As ListObject) As Boolean
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim connNameAlt As String
    Dim success As Boolean
    
    ' Initialize
    Solution5_ConnectionStateReset = False
    connName = "pgGet510kData"
    connNameAlt = "Query - " & connName
    
    ' Reset environment
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    DoEvents
    
    ' Find the connection
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt)
    
    If wbConn Is Nothing Then
        MsgBox "Could not find Power Query connection to refresh.", vbExclamation, "Refresh Error"
        GoTo Cleanup
    End If
    
    ' Try to reset connection state
    Call CleanupDuplicateConnections ' Call to the existing cleanup function
    
    ' Force EnableRefresh if available
    On Error Resume Next
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
        wbConn.OLEDBConnection.EnableRefresh = True
    End If
    
    ' Perform the refresh with cleaner state
    On Error Resume Next
    wbConn.Refresh
    success = (Err.Number = 0)
    
    If Not success Then
        Debug.Print "Connection state reset failed: " & Err.Description
    End If
    
Cleanup:
    ' Restore environment
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set wbConn = Nothing
    Solution5_ConnectionStateReset = success
    
End Function

' This is a placeholder reference to the existing connection cleanup function
' from mod_DataIO module that should already exist in the project
Private Sub CleanupDuplicateConnections()
    ' This should call the existing function in mod_DataIO
    mod_DataIO.CleanupDuplicateConnections
End Sub

