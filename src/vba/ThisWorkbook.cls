Option Explicit
Attribute VB_Name = "ThisWorkbook"

' ==========================================================================
' ===                  THISWORKBOOK MODULE CODE                          ===
' ==========================================================================
' This code handles the automatic workbook events:
' - Workbook_Open: Entry point for the automated 510(k) processing
' - Workbook_BeforeClose: Performs cleanup operations like trimming the log

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
    ' Assumes RefreshPowerQuery is a Public Function in mod_510k_Processor
    ' It now requires the ListObject (table) as input.
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim refreshSuccess As Boolean: refreshSuccess = False
    Const DATA_SHEET_NAME_LOCAL As String = "CurrentMonthData" ' Use local literal

    On Error Resume Next ' Handle potential errors getting sheet/table
    ' Use the local literal sheet name to avoid cross-module reference issues
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME_LOCAL)
    If Not wsData Is Nothing Then
        Set tblData = wsData.ListObjects(1) ' Assuming the first table is the target
    End If
    On Error GoTo OpenErrorHandler ' Restore error handler

    If tblData Is Nothing Then
        ' Use the local literal here too for the error message
        Debug.Print Time & " - ERROR: Could not find data table on sheet '" & DATA_SHEET_NAME_LOCAL & "' during Workbook_Open refresh."
        MsgBox "Error: Could not find the data table on sheet '" & DATA_SHEET_NAME_LOCAL & "'." & vbCrLf & _
               "Cannot refresh Power Query.", vbCritical, "Table Not Found"
        ' Decide if you want to stop here or continue without refresh
        ' GoTo OpenCleanExit ' Option to stop
    Else
        ' Call RefreshPowerQuery with the table object
        refreshSuccess = mod_DataIO.RefreshPowerQuery(tblData) ' Corrected module reference
        If refreshSuccess Then
            Debug.Print Time & " - Power Query refresh step completed for table '" & tblData.Name & "'."
        Else
            Debug.Print Time & " - Power Query refresh failed or was skipped for table '" & tblData.Name & "'. See previous logs."
            ' Critical Decision: Should we stop if refresh fails?
            MsgBox "Warning: Failed to refresh data from the FDA source." & vbCrLf & _
                   "Processing will continue with existing data, which may be outdated.", vbExclamation, "Data Refresh Issue"
            ' Uncomment the next line to completely stop if refresh fails:
            ' GoTo OpenCleanExit
        End If
    End If

    ' Clean up objects used just for refresh call
    Set tblData = Nothing
    Set wsData = Nothing

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

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Purpose: Performs any necessary cleanup operations before the workbook is closed
    '          1. Trims the RunLog sheet to prevent excessive growth over time
    '          2. [Optional] Could add other cleanup tasks here as needed
    
    On Error Resume Next ' Handle potential errors gracefully - we don't want to prevent closing
    
    ' --- 1. Trim the Log Sheet ---
    mod_Logger.TrimRunLog 5000 ' Keep the most recent 5000 rows in the log, delete older entries
    
    ' --- 2. [Optional] Any other cleanup tasks could be added here ---
    ' Examples:
    ' - Ensure all sheets are properly archived
    ' - Save user preferences/settings
    ' - Clean up temp files/data
    
    On Error GoTo 0 ' Restore normal error handling
    
    ' Always let workbook close by not setting Cancel = True
End Sub
