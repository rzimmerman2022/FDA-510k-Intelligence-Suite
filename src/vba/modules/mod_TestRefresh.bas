Attribute VB_Name = "mod_TestRefresh"
Option Explicit

' --- Test Option Constants ---
Private Const OPT_REFRESH_WITH_BACKGROUND_FALSE As Boolean = True  ' Set to False to test with BackgroundQuery=True
Private Const OPT_REFRESH_WITH_CALCULATION_MANUAL As Boolean = True ' Set to False to test with current calculation mode

Sub TestRefreshOnly()
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim connNameAlt1 As String
    Dim connNameAlt2 As String
    Dim found As Boolean

    found = False

    ' --- Common Power Query Connection Naming Patterns ---
    ' *** Adjust these based on your primary query name if it's not pgGet510kData ***
    Const baseQueryName As String = "pgGet510kData" 

    connName = baseQueryName                    ' e.g., pgGet510kData
    connNameAlt1 = "Query - " & baseQueryName  ' e.g., Query - pgGet510kData
    connNameAlt2 = "Connection " & baseQueryName ' Less common, but possible

    ' Add more patterns here if needed

    Debug.Print "TestRefreshOnly: Searching for connection..."

    ' --- Find the Workbook Connection using multiple patterns ---
    On Error Resume Next ' Try finding connection by various names
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt1)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt2)
    ' Add more checks if other naming patterns exist
    On Error GoTo 0 ' Restore default error handling

    If wbConn Is Nothing Then
        MsgBox "Test Error: Could not find connection using patterns like '" & connName & "' or '" & connNameAlt1 & "'. Please check Data > Queries & Connections > Connections tab for the exact primary connection name.", vbCritical, "Connection Not Found"
        Exit Sub
    Else
        Debug.Print "TestRefreshOnly: Found connection: '" & wbConn.Name & "'"
        found = True
    End If

    MsgBox "Found connection: '" & wbConn.Name & "'" & vbCrLf & vbCrLf & "Attempting to refresh...", vbInformation, "Test Refresh"

    ' --- Attempt Refresh ---
    On Error Resume Next ' Isolate error specifically on the refresh line

    ' Ensure BackgroundQuery is False (Important!)
    wbConn.OLEDBConnection.BackgroundQuery = False ' Set before refreshing
    If Err.Number <> 0 Then
         Debug.Print "TestRefreshOnly: Error setting BackgroundQuery=False for '" & wbConn.Name & "'. Err: " & Err.Description
         Err.Clear ' Clear error and try refresh anyway
    End If

    ' The actual refresh call
    wbConn.Refresh

    ' --- Check Result ---
    If Err.Number <> 0 Then
        MsgBox "TEST FAILED refreshing connection '" & wbConn.Name & "'." & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbCritical, "Isolated Refresh Error"
        Debug.Print "TestRefreshOnly: FAILED refreshing connection '" & wbConn.Name & "'. Error " & Err.Number & ": " & Err.Description
    Else
        MsgBox "TEST SUCCEEDED: Refresh command completed for connection '" & wbConn.Name & "'.", vbInformation, "Isolated Refresh Success"
        Debug.Print "TestRefreshOnly: SUCCEEDED refreshing connection '" & wbConn.Name & "'."
    End If
    On Error GoTo 0 ' Restore default error handling

    Set wbConn = Nothing
End Sub

Sub TestRefreshWithOptions()
    ' Purpose: More comprehensive test for diagnosing refresh issues with various settings
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim connNameAlt1 As String
    Dim connNameAlt2 As String
    Dim found As Boolean
    Dim originalCalc As XlCalculation
    Dim debugMsg As String

    found = False
    originalCalc = Application.Calculation ' Store current calculation mode
    debugMsg = "TestRefreshWithOptions:" & vbCrLf & "-------------------" & vbCrLf

    ' --- Common Power Query Connection Naming Patterns ---
    Const baseQueryName As String = "pgGet510kData" 

    connName = baseQueryName
    connNameAlt1 = "Query - " & baseQueryName
    connNameAlt2 = "Connection " & baseQueryName

    debugMsg = debugMsg & "Looking for connection with patterns: '" & connName & "', '" & connNameAlt1 & "'" & vbCrLf
    Debug.Print "TestRefreshWithOptions: Searching for connection..."

    ' --- Find the Workbook Connection using multiple patterns ---
    On Error Resume Next ' Try finding connection by various names
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt1)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt2)
    On Error GoTo 0 ' Restore default error handling

    If wbConn Is Nothing Then
        debugMsg = debugMsg & "ERROR: Could not find connection!" & vbCrLf
        MsgBox "Test Error: Could not find connection using patterns like '" & connName & "' or '" & connNameAlt1 & "'." & vbCrLf & _
               "Please check Data > Queries & Connections > Connections tab for the exact primary connection name.", vbCritical, "Connection Not Found"
        Debug.Print debugMsg
        Exit Sub
    Else
        debugMsg = debugMsg & "Found connection: '" & wbConn.Name & "'" & vbCrLf
        Debug.Print "TestRefreshWithOptions: Found connection: '" & wbConn.Name & "'"
        found = True
    End If

    ' --- Display initial settings ---
    debugMsg = debugMsg & "Current calculation mode: " & Application.Calculation & vbCrLf
    debugMsg = debugMsg & "OPT_REFRESH_WITH_CALCULATION_MANUAL = " & OPT_REFRESH_WITH_CALCULATION_MANUAL & vbCrLf

    ' Check if it has OLEDBConnection
    On Error Resume Next
    Dim hasOLEDB As Boolean: hasOLEDB = Not wbConn.OLEDBConnection Is Nothing
    Dim currBackgroundQuery As Boolean
    
    If hasOLEDB Then
        currBackgroundQuery = wbConn.OLEDBConnection.BackgroundQuery
        debugMsg = debugMsg & "Connection has OLEDBConnection" & vbCrLf
        debugMsg = debugMsg & "Current BackgroundQuery = " & currBackgroundQuery & vbCrLf
        debugMsg = debugMsg & "OPT_REFRESH_WITH_BACKGROUND_FALSE = " & OPT_REFRESH_WITH_BACKGROUND_FALSE & vbCrLf
    Else
        debugMsg = debugMsg & "Connection does NOT have OLEDBConnection" & vbCrLf
    End If
    On Error GoTo 0
    
    ' User prompt to continue
    If MsgBox("Found connection: '" & wbConn.Name & "'" & vbCrLf & vbCrLf & _
              "Test will use the following settings:" & vbCrLf & _
              "- Set calculation to manual during refresh: " & OPT_REFRESH_WITH_CALCULATION_MANUAL & vbCrLf & _
              "- Set BackgroundQuery to False: " & OPT_REFRESH_WITH_BACKGROUND_FALSE & vbCrLf & vbCrLf & _
              "Continue with refresh attempt?", vbQuestion + vbYesNo, "Test Refresh With Options") = vbNo Then
        Debug.Print debugMsg & "User cancelled test."
        Exit Sub
    End If

    ' --- Prepare environment for refresh ---
    On Error Resume Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Set calculation mode if option enabled
    If OPT_REFRESH_WITH_CALCULATION_MANUAL Then
        Application.Calculation = xlCalculationManual
        debugMsg = debugMsg & "Set calculation to MANUAL" & vbCrLf
    End If
    
    ' Set BackgroundQuery if it has OLEDBConnection and option enabled
    If hasOLEDB And OPT_REFRESH_WITH_BACKGROUND_FALSE Then
        wbConn.OLEDBConnection.BackgroundQuery = False
        If Err.Number <> 0 Then
            debugMsg = debugMsg & "ERROR setting BackgroundQuery=False: " & Err.Description & vbCrLf
            Err.Clear
        Else
            debugMsg = debugMsg & "Successfully set BackgroundQuery = False" & vbCrLf
        End If
    End If
    
    ' Allow Excel to process events before refresh
    DoEvents
    debugMsg = debugMsg & "DoEvents called before refresh" & vbCrLf
    
    ' --- The actual refresh attempt ---
    debugMsg = debugMsg & "Attempting refresh..." & vbCrLf
    wbConn.Refresh ' The core command we're testing

    ' --- Check Result ---
    If Err.Number <> 0 Then
        debugMsg = debugMsg & "REFRESH FAILED with error " & Err.Number & ": " & Err.Description & vbCrLf
        MsgBox "TEST FAILED refreshing connection '" & wbConn.Name & "'." & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "Check the Immediate window (Ctrl+G in VBA Editor) for detailed diagnostic info.", vbCritical, "Isolated Refresh Error"
    Else
        debugMsg = debugMsg & "REFRESH SUCCEEDED!" & vbCrLf
        MsgBox "TEST SUCCEEDED: Refresh command completed without error for connection '" & wbConn.Name & "'." & vbCrLf & vbCrLf & _
               "Check the Immediate window (Ctrl+G in VBA Editor) for detailed diagnostic info.", vbInformation, "Isolated Refresh Success"
    End If
    On Error GoTo 0 ' Restore default error handling

    ' --- Cleanup ---
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    debugMsg = debugMsg & "Environment restored. Test complete." & vbCrLf
    
    Debug.Print debugMsg
    Set wbConn = Nothing
End Sub

Sub TestAllOptions()
    ' Purpose: Systematically test refresh with different combinations of settings
    Dim results(1 To 4) As String
    Dim wbConn As WorkbookConnection
    Dim connName As String
    Dim connNameAlt1 As String
    Dim connNameAlt2 As String
    Dim originalCalc As XlCalculation
    Dim testNum As Integer
    Dim currTest As String
    Dim errMsg As String
    
    originalCalc = Application.Calculation
    testNum = 0
    
    ' --- Find the connection first ---
    connName = "pgGet510kData"
    connNameAlt1 = "Query - " & connName
    connNameAlt2 = "Connection " & connName
    
    On Error Resume Next
    Set wbConn = ThisWorkbook.Connections(connName)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt1)
    If wbConn Is Nothing Then Set wbConn = ThisWorkbook.Connections(connNameAlt2)
    On Error GoTo 0
    
    If wbConn Is Nothing Then
        MsgBox "Could not find the Power Query connection. Cannot run tests.", vbCritical, "Test Failed"
        Exit Sub
    End If
    
    If MsgBox("This test will systematically try refreshing with 4 different setting combinations." & vbCrLf & _
              "Each test will display a message with the results." & vbCrLf & vbCrLf & _
              "Continue with the test series?", vbQuestion + vbYesNo, "Systematic Testing") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' --- TEST 1: Background=False, Calculation=Manual ---
    testNum = 1
    currTest = "TEST " & testNum & ": BackgroundQuery=False, Calculation=Manual"
    Debug.Print currTest
    
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    wbConn.Refresh
    
    If Err.Number <> 0 Then
        errMsg = Err.Description
        results(testNum) = "FAILED - Error: " & errMsg
    Else
        results(testNum) = "SUCCEEDED"
    End If
    Err.Clear
    On Error GoTo 0
    
    ' --- TEST 2: Background=False, Calculation=Automatic ---
    testNum = 2
    currTest = "TEST " & testNum & ": BackgroundQuery=False, Calculation=Automatic"
    Debug.Print currTest
    
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = False
    End If
    wbConn.Refresh
    
    If Err.Number <> 0 Then
        errMsg = Err.Description
        results(testNum) = "FAILED - Error: " & errMsg
    Else
        results(testNum) = "SUCCEEDED"
    End If
    Err.Clear
    On Error GoTo 0
    
    ' --- TEST 3: Background=True, Calculation=Manual ---
    testNum = 3
    currTest = "TEST " & testNum & ": BackgroundQuery=True, Calculation=Manual"
    Debug.Print currTest
    
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = True
    End If
    wbConn.Refresh
    
    If Err.Number <> 0 Then
        errMsg = Err.Description
        results(testNum) = "FAILED - Error: " & errMsg
    Else
        results(testNum) = "SUCCEEDED"
    End If
    Err.Clear
    On Error GoTo 0
    
    ' --- TEST 4: Background=True, Calculation=Automatic ---
    testNum = 4
    currTest = "TEST " & testNum & ": BackgroundQuery=True, Calculation=Automatic"
    Debug.Print currTest
    
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    If Not wbConn.OLEDBConnection Is Nothing Then
        wbConn.OLEDBConnection.BackgroundQuery = True
    End If
    wbConn.Refresh
    
    If Err.Number <> 0 Then
        errMsg = Err.Description
        results(testNum) = "FAILED - Error: " & errMsg
    Else
        results(testNum) = "SUCCEEDED"
    End If
    Err.Clear
    On Error GoTo 0
    
    ' --- Cleanup and Report Results ---
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Test Results Summary:" & vbCrLf & vbCrLf & _
           "TEST 1: BackgroundQuery=False, Calculation=Manual - " & results(1) & vbCrLf & _
           "TEST 2: BackgroundQuery=False, Calculation=Automatic - " & results(2) & vbCrLf & _
           "TEST 3: BackgroundQuery=True, Calculation=Manual - " & results(3) & vbCrLf & _
           "TEST 4: BackgroundQuery=True, Calculation=Automatic - " & results(4) & vbCrLf & vbCrLf & _
           "Check the Immediate window for more details.", vbInformation, "Systematic Test Results"
    
    Debug.Print "TEST RESULTS SUMMARY:"
    Debug.Print "-------------------"
    Debug.Print "TEST 1: BackgroundQuery=False, Calculation=Manual - " & results(1)
    Debug.Print "TEST 2: BackgroundQuery=False, Calculation=Automatic - " & results(2)
    Debug.Print "TEST 3: BackgroundQuery=True, Calculation=Manual - " & results(3)
    Debug.Print "TEST 4: BackgroundQuery=True, Calculation=Automatic - " & results(4)
    
    Set wbConn = Nothing
End Sub
