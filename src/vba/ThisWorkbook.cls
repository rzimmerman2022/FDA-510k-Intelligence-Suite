Option Explicit

' ==========================================================================
' ===                  THISWORKBOOK MODULE CODE                          ===
' ==========================================================================
' --- This code runs automatically when the workbook is opened ---

Private Sub Workbook_Open()
    ' Purpose: Entry point for the automated 510(k) processing.
    '          Refreshes data on every open via Power Query, then calls
    '          the main processing module which handles scoring, formatting,
    '          and conditional archiving.

    Dim procStartTime As Double: procStartTime = Timer ' For timing measurement

    ' --- Basic error handling for the open event ---
    On Error GoTo OpenErrorHandler

    ' --- 1. Check if Workbook is ReadOnly ---
    If Me.ReadOnly Then
        Debug.Print Time & " - Workbook opened ReadOnly. Auto-processing skipped."
        Exit Sub ' Do not proceed if the workbook cannot be modified
    End If

    ' --- 2. Prepare Application State for Performance ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False ' Prevent other event triggers during run
    Application.Calculation = xlCalculationManual ' Prevent calculation delays
    Application.Cursor = xlWait ' Indicate busy state to user

    ' --- 3. Always Refresh Power Query Data First ---
    Application.StatusBar = "Refreshing FDA data (Step 1/2)..."
    Debug.Print Time & " - Workbook_Open: Calling RefreshPowerQuery..."

    ' Call the refresh function from the other module.
    ' Assumes PQ_CONNECTION_NAME is a Public Const in mod_510k_Processor
    ' Assumes RefreshPowerQuery is a Public Function in mod_510k_Processor or this module
    If mod_510k_Processor.RefreshPowerQuery(mod_510k_Processor.PQ_CONNECTION_NAME) Then
         Debug.Print Time & " - Power Query refresh step completed."
    Else
         Debug.Print Time & " - Power Query refresh failed or was skipped. See previous logs."
         ' Critical Decision: Should we stop if refresh fails?
         MsgBox "Warning: Failed to refresh data from the FDA source." & vbCrLf & _
                "Processing will continue with existing data, which may be outdated.", vbExclamation, "Data Refresh Issue"
         ' Uncomment the next line to completely stop if refresh fails:
         ' GoTo OpenCleanExit
    End If

    ' --- 4. Call the Main Processing Sub (Scoring, Formatting, Archiving) ---
    '    This sub contains the Day Guard logic.
    Application.StatusBar = "Processing leads and finalizing sheet (Step 2/2)..."
    Debug.Print Time & " - Workbook_Open: Calling ProcessMonthly510k..."

    Call mod_510k_Processor.ProcessMonthly510k ' Call the main routine in the other module

    ' --- 5. Final Status Update ---
    '    Note: A more detailed completion message is shown at the end of ProcessMonthly510k
    Application.StatusBar = "Workbook ready."
    Debug.Print Time & " - Workbook_Open sequence completed in " & Format(Timer - procStartTime, "0.0") & " seconds."


' --- Clean Exit Point (for normal completion or after handled error) ---
OpenCleanExit:
    ' --- Restore Application Settings ---
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False ' Clear status bar message
    Application.Cursor = xlDefault ' Restore default cursor
    Exit Sub ' End the Workbook_Open procedure


' --- Error Handler for Workbook_Open ---
OpenErrorHandler:
    MsgBox "A critical error occurred during the workbook opening sequence: " & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & vbCrLf & _
           "Attempting to restore settings.", vbCritical, "Workbook Open Error"
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Debug.Print Time & " - CRITICAL ERROR in Workbook_Open: " & Err.Description & " (Error No. " & Err.Number & ")"
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    ' Attempt to jump to the cleanup routine to restore settings
    Resume OpenCleanExit

End Sub ' End of Workbook_Open

