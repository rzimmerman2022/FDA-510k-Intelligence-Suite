Attribute VB_Name = "mod_TestWithContext"
Option Explicit

' ==========================================================================
' Module      : mod_TestWithContext
' Author      : Cline (AI Assistant)
' Date        : May 7, 2025
' Purpose     : Advanced diagnostic tests for Power Query refresh issues that 
'               specifically target context-related issues that might be 
'               causing refresh to fail when triggered automatically but
'               succeed when triggered manually.
' ==========================================================================

Public Enum RefreshTriggerContext
    rtcDirectCall = 1           ' Direct Sub call
    rtcCommandBarButton = 2     ' Via CommandBar button click
    rtcYesNoPrompt = 3          ' Via MsgBox Yes/No response
    rtcTimerEvent = 4           ' Via Application.OnTime event
End Enum

' --- Test Variables ---
Private mTestRunning As Boolean   ' Flag to prevent multiple test runs
Private mConnectionName As String ' Name of the found connection
Private mTestTimer As Double      ' For timing delayed tests

Sub TestContextDependentRefresh()
    ' Entry point - tests refresh in different execution contexts
    Dim wbConn As WorkbookConnection
    
    ' Find the connection first
    If Not FindPowerQueryConnection(wbConn) Then Exit Sub
    
    ' Store the connection name for other tests to use
    mConnectionName = wbConn.Name
    Set wbConn = Nothing
    
    ' Prevent multiple test runs
    If mTestRunning Then
        MsgBox "A test is already running. Please wait for it to complete.", vbExclamation, "Test In Progress"
        Exit Sub
    End If
    
    ' Start the test sequence
    If MsgBox("This test will try refreshing the connection in different execution contexts to isolate why" & vbCrLf & _
              "automatic refresh fails but manual refresh works." & vbCrLf & vbCrLf & _
              "The test will:" & vbCrLf & _
              "1. Create a temporary CommandBar button" & vbCrLf & _
              "2. Show dialog boxes you'll need to respond to" & vbCrLf & _
              "3. Run a delayed timer event" & vbCrLf & vbCrLf & _
              "Continue with the tests?", vbQuestion + vbYesNo, "Context-Dependent Refresh Test") = vbNo Then
        Exit Sub
    End If
    
    mTestRunning = True
    
    ' Start with a direct call test
    MsgBox "First we'll try a direct refresh call from this procedure.", vbInformation, "Direct Call Test"
    TestSpecificContext rtcDirectCall
    
    ' Now set up the command bar test
    MsgBox "Next we'll create a temporary toolbar button." & vbCrLf & vbCrLf & _
           "Please click the 'Test Refresh' button that will appear in the toolbar.", vbInformation, "CommandBar Test"
    CreateTemporaryTestButton
    
    ' The message box test will be triggered by the direct test function
    ' The timer test will be scheduled at the end of the test sequence
End Sub

Private Function FindPowerQueryConnection(ByRef wbConn As WorkbookConnection) As Boolean
    ' Helper function to find the Power Query connection
    Dim connName As String
    Dim connNameAlt1 As String
    Dim connNameAlt2 As String
    
    ' Reset output
    Set wbConn = Nothing
    FindPowerQueryConnection = False
    
    ' Common naming patterns
    connName = "pgGet510kData"
    connNameAlt1 = "Query - " & connName
    connNameAlt2 = "Connection " & connName
    
    Debug.Print "FindPowerQueryConnection: Searching for connection..."
    
    ' Try finding connection by various names
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt1)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt2)
    On Error GoTo 0
    
    If wbConn Is Nothing Then
        MsgBox "Test Error: Could not find connection using patterns like '" & connName & "' or '" & connNameAlt1 & "'." & vbCrLf & _
               "Please check Data > Queries & Connections > Connections tab for the exact primary connection name.", vbCritical, "Connection Not Found"
        Debug.Print "FindPowerQueryConnection: Failed to find connection"
        FindPowerQueryConnection = False
    Else
        Debug.Print "FindPowerQueryConnection: Found connection: '" & wbConn.Name & "'"
        FindPowerQueryConnection = True
    End If
End Function

Sub TestSpecificContext(context As RefreshTriggerContext)
    ' Executes a refresh test in a specific context
    Dim wbConn As WorkbookConnection
    Dim contextName As String
    Dim success As Boolean
    
    ' Get context name for reporting
    Select Case context
        Case rtcDirectCall: contextName = "Direct VBA Call"
        Case rtcCommandBarButton: contextName = "CommandBar Button Click"
        Case rtcYesNoPrompt: contextName = "Yes/No Dialog Response"
        Case rtcTimerEvent: contextName = "Timer Event"
        Case Else: contextName = "Unknown Context"
    End Select
    
    Debug.Print "TestSpecificContext: Starting test in context: " & contextName
    
    ' Find connection
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(mConnectionName)
    If wbConn Is Nothing Then
        If Not FindPowerQueryConnection(wbConn) Then Exit Sub
        mConnectionName = wbConn.Name
    End If
    On Error GoTo 0
    
    ' --- Perform the refresh ---
    Debug.Print "TestSpecificContext: Attempting refresh in context: " & contextName
    success = RefreshInSpecificContext(wbConn, context)
    
    ' --- Handle next steps based on the current context ---
    Select Case context
        Case rtcDirectCall
            ' After direct call, ask if the user wants to try the Yes/No prompt test
            If MsgBox("Direct call test " & IIf(success, "SUCCEEDED", "FAILED") & "." & vbCrLf & vbCrLf & _
                       "Would you like to test refreshing via Yes/No prompt response now?", vbQuestion + vbYesNo, "Try Yes/No Test") = vbYes Then
                TestSpecificContext rtcYesNoPrompt
            End If
            
        Case rtcCommandBarButton
            ' After command bar test, schedule the timer test
            RemoveTemporaryTestButton
            MsgBox "CommandBar button test " & IIf(success, "SUCCEEDED", "FAILED") & "." & vbCrLf & vbCrLf & _
                   "Next we'll test a refresh triggered by a timer event." & vbCrLf & _
                   "The refresh will happen in 5 seconds after you click OK.", vbInformation, "Timer Test"
            ScheduleTimerTest
            
        Case rtcYesNoPrompt
            ' After Yes/No test, remind about the command bar button
            MsgBox "Yes/No prompt test " & IIf(success, "SUCCEEDED", "FAILED") & "." & vbCrLf & vbCrLf & _
                   "Remember to click the 'Test Refresh' button in the toolbar to continue testing.", vbInformation, "Continue Testing"
            
        Case rtcTimerEvent
            ' After timer test, show final results and clean up
            MsgBox "Timer event test " & IIf(success, "SUCCEEDED", "FAILED") & "." & vbCrLf & vbCrLf & _
                   "All context tests are now complete.", vbInformation, "Tests Complete"
            CompleteTestSequence
    End Select
    
    Set wbConn = Nothing
End Sub

Private Function RefreshInSpecificContext(wbConn As WorkbookConnection, context As RefreshTriggerContext) As Boolean
    ' Performs the actual refresh operation in the specified context
    Dim contextName As String
    Dim originalCalc As XlCalculation
    
    ' Default to failure
    RefreshInSpecificContext = False
    
    ' Get context name for debug output
    Select Case context
        Case rtcDirectCall: contextName = "Direct VBA Call"
        Case rtcCommandBarButton: contextName = "CommandBar Button Click"
        Case rtcYesNoPrompt: contextName = "Yes/No Dialog Response"
        Case rtcTimerEvent: contextName = "Timer Event"
        Case Else: contextName = "Unknown Context"
    End Select
    
    On Error Resume Next
    
    ' Store original calculation mode and apply settings
    originalCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Prepare the connection if it's OLEDB
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    
    Debug.Print "RefreshInSpecificContext: Attempting refresh in context '" & contextName & "'"
    
    ' Execute DoEvents to flush the message queue before refresh
    DoEvents
    
    ' Perform the refresh
    wbConn.Refresh
    
    ' Check for errors
    If Err.Number <> 0 Then
        Debug.Print "RefreshInSpecificContext: FAILED in context '" & contextName & "'. Error " & Err.Number & ": " & Err.Description
        MsgBox "Refresh FAILED in context: " & contextName & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbCritical, "Context Test Failed"
        RefreshInSpecificContext = False
    Else
        Debug.Print "RefreshInSpecificContext: SUCCEEDED in context '" & contextName & "'"
        MsgBox "Refresh SUCCEEDED in context: " & contextName, vbInformation, "Context Test Succeeded"
        RefreshInSpecificContext = True
    End If
    
    ' Restore original settings
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    On Error GoTo 0
End Function

Private Sub CreateTemporaryTestButton()
    ' Creates a temporary CommandBar button for testing button-click refresh
    Dim testBar As CommandBar
    Dim testButton As CommandBarButton
    
    On Error Resume Next
    
    ' First try to remove any existing test bar
    Application.CommandBars("RefreshTest").Delete
    
    ' Create a new CommandBar
    Set testBar = Application.CommandBars.Add(Name:="RefreshTest", Position:=msoBarTop, Temporary:=True)
    testBar.Visible = True
    
    ' Add a button to the CommandBar
    Set testButton = testBar.Controls.Add(Type:=msoControlButton)
    With testButton
        .Caption = "Test Refresh"
        .Style = msoButtonCaption
        .FaceId = 59 ' Refresh icon
        .OnAction = "'" & ThisWorkbook.Name & "'!mod_TestWithContext.CommandBarRefreshTest"
    End With
    
    On Error GoTo 0
End Sub

Sub CommandBarRefreshTest()
    ' Called when the CommandBar button is clicked
    TestSpecificContext rtcCommandBarButton
End Sub

Private Sub RemoveTemporaryTestButton()
    ' Removes the temporary CommandBar
    On Error Resume Next
    Application.CommandBars("RefreshTest").Delete
    On Error GoTo 0
End Sub

Private Sub ScheduleTimerTest()
    ' Schedules a timer-based refresh test
    mTestTimer = Now + TimeSerial(0, 0, 5) ' 5 seconds from now
    Application.OnTime mTestTimer, "'" & ThisWorkbook.Name & "'!mod_TestWithContext.TimerRefreshTest"
End Sub

Sub TimerRefreshTest()
    ' Called by the Application.OnTime event
    TestSpecificContext rtcTimerEvent
End Sub

Private Sub CompleteTestSequence()
    ' Clean up after all tests are complete
    RemoveTemporaryTestButton
    mTestRunning = False
    
    ' Cancel any pending timer if it exists
    On Error Resume Next
    If mTestTimer <> 0 Then
        Application.OnTime mTestTimer, "'" & ThisWorkbook.Name & "'!mod_TestWithContext.TimerRefreshTest", , False
    End If
    On Error GoTo 0
    
    Debug.Print "TestWithContext: All tests completed"
End Sub

Sub CancelAllTests()
    ' Emergency cleanup - can be called to cancel all tests
    CompleteTestSequence
    MsgBox "All context tests have been cancelled.", vbInformation, "Tests Cancelled"
End Sub
