' ==========================================================================
' Module      : ThisWorkbook (Class Module)
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This is the code module associated with the Excel Workbook
'               object itself. It primarily contains event handler procedures
'               that respond to workbook-level events.
'
' Key Procedures:
'               - Workbook_Open(): Automatically executes when the workbook is
'                 opened. Auto-update functionality (refresh, processing)
'                 is currently COMMENTED OUT within this sub.
'               - Workbook_BeforeClose(Cancel As Boolean): Executes just before
'                 the workbook closes. Currently used to perform cleanup tasks,
'                 specifically flushing the log buffer and trimming the RunLog
'                 sheet using mod_Logger.
'
' Dependencies: - mod_DataIO: Called by Workbook_Open (currently commented out).
'               - mod_510k_Processor: Called by Workbook_Open (currently commented out).
'               - mod_Logger: Calls FlushLogBuf and TrimRunLog.
'               - Implicitly relies on constants/settings defined in mod_Config
'                 via the modules it calls.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-05-08  Cline (AI)      - Improved error handling in Workbook_Open to ensure events
'                               are always re-enabled even if a critical error occurs
'                               - Added global variable g_TargetMonth to avoid month logic duplication
' 2025-05-01  Cline (AI)      - Commented out auto-update calls within Workbook_Open
'                               (RefreshPowerQuery, ProcessMonthly510k) per user request.
' 2025-05-01  Cline (AI)      - Removed previous commenting out of entire Workbook_Open.
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-29  Cline (AI)      - Corrected module reference for RefreshPowerQuery call
'                               (changed from mod_510k_Processor to mod_DataIO).
' 2025-04-30  Cline (AI)      - Removed explicit module qualifier from RefreshPowerQuery
'                               call to attempt resolution of compile error.
' 2025-04-30  Cline (AI)      - Reverted: Restored explicit module qualifier (mod_DataIO.)
'                               for RefreshPowerQuery call.
' 2025-04-30  Cline (AI)      - Added call to mod_DataIO.TestCall for debugging compile error.
' 2025-04-30  Cline (AI)      - Removed Debug.Print statements previously added for tracing.
' 2025-04-30  Cline (AI)      - Added call to mod_Logger.FlushLogBuf in Workbook_BeforeClose
'                               to ensure logs are written before closing/trimming.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit
' This code handles the automatic workbook events:
' - Workbook_Open: Entry point for the automated 510(k) processing
' - Workbook_BeforeClose: Performs cleanup operations like trimming the log

' Global variable to avoid month determination duplication
' Set here by DetermineReportingMonth and used by ProcessMonthly510k
Public g_TargetMonth As Date

Private Sub Workbook_Open()
    ' Purpose: Entry point for the automated 510(k) processing.
    '          NEW FUNCTIONALITY: 
    '          1. Check if today's date is on/after the 10th of the month
    '          2. Determine which period to process (previous month if after 10th)
    '          3. Check if we've already created the archive sheet for this period
    '          4. If needed, run a synchronous PQ refresh and process the data

    Dim procStartTime As Double: procStartTime = Timer ' For timing measurement
    Dim origEnableEvents As Boolean: origEnableEvents = Application.EnableEvents ' Store original state

    ' --- Put error handling FIRST to ensure events are always re-enabled ---
    ' This is critical: Even if a panic error occurs, we need to restore application state
    On Error GoTo OpenErrorHandler

    ' --- Disable events to prevent Power Query from potentially re-triggering Workbook_Open ---
    Application.EnableEvents = False
    
    ' --- 1. Check if Workbook is ReadOnly ---
    If Me.ReadOnly Then
        Debug.Print Time & " - Workbook opened ReadOnly. Auto-processing skipped."
        GoTo OpenCleanExit ' Use clean exit to restore settings
    End If
    
    ' --- NEW CODE: Determine which period to process based on date ---
    Dim tgtMonthFirst As Date, tgtSheetName As String
    
    ' 1) Decide which period to process based on the current date
    tgtMonthFirst = DetermineReportingMonth(Date)
    
    ' Store in global variable for ProcessMonthly510k to use
    ' This avoids duplicate month determination logic
    g_TargetMonth = tgtMonthFirst
    
    tgtSheetName = Format(tgtMonthFirst, "mmm-yyyy") ' e.g., "Apr-2025"
    
    ' 2) If we've already built the archive sheet for this period, bail out
    If SheetExists(tgtSheetName) Then
        Application.StatusBar = "Month " & tgtSheetName & " already archived. No processing needed."
        Debug.Print Time & " - Month " & tgtSheetName & " already archived. Auto-processing skipped."
        GoTo OpenCleanExit
    End If
    
    ' 3) We need to process this month - prepare application state
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
    
    ' 4) Force a synchronous PQ refresh
    Application.StatusBar = "Refreshing FDA data for month " & tgtSheetName & "..."
    If Not Refresh510kConnection() Then
        MsgBox "Warning: Power Query refresh failed. Will attempt to continue with existing data.", vbExclamation, "Refresh Warning"
    End If
    
    ' 5) Run the full processing pipeline with the target month as parameter
    Application.StatusBar = "Processing data for " & tgtSheetName & "..."
    Debug.Print Time & " - Auto-processing " & tgtSheetName & " based on date rule (current date: " & Format(Date, "mm/dd/yyyy") & ")"
    
    ' NOTE: Using g_TargetMonth global variable to share the target month
    ' ProcessMonthly510k should check g_TargetMonth and use it if not Empty/zero
    ' Future enhancement: Modify ProcessMonthly510k to accept a date parameter directly
    Call mod_510k_Processor.ProcessMonthly510k
    
    ' 6) Final success message
    Application.StatusBar = "Processing complete for " & tgtSheetName
    Debug.Print Time & " - Workbook_Open auto-processing completed in " & Format(Timer - procStartTime, "0.0") & " seconds."

' --- FINALLY block - this ensures application state is always restored ---
OpenCleanExit:
    ' --- Restore Application Settings (Must execute even after errors) ---
    Application.ScreenUpdating = True
    Application.EnableEvents = origEnableEvents ' Restore original EnableEvents state
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
    ' Jump to our FINALLY block to ensure settings are restored
    ' The Resume ensures we'll continue with the clean-exit code
    ' no matter what kind of error happened
    Resume OpenCleanExit

End Sub ' End of Workbook_Open

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Purpose: Performs any necessary cleanup operations before the workbook is closed
    '          1. Flushes the log buffer
    '          2. Trims the RunLog sheet to prevent excessive growth over time
    '          3. [Optional] Could add other cleanup tasks here as needed

    On Error Resume Next ' Handle potential errors gracefully - we don't want to prevent closing

    ' --- 1. Flush Log Buffer ---
    ' Ensure any remaining logs in memory are written to the sheet before closing/trimming.
    mod_Logger.FlushLogBuf

    ' --- 2. Trim the Log Sheet ---
    mod_Logger.TrimRunLog 5000 ' Keep the most recent 5000 rows in the log, delete older entries

    ' --- 3. [Optional] Any other cleanup tasks could be added here ---
    ' Examples:
    ' - Ensure all sheets are properly archived
    ' - Save user preferences/settings
    ' - Clean up temp files/data

    On Error GoTo 0 ' Restore normal error handling

    ' Always let workbook close by not setting Cancel = True
End Sub

' === NEW HELPER FUNCTIONS ===

' Function to determine which reporting month to process based on the current date
' Rule: If day >= 10, process previous month; Otherwise, process month before previous
Public Function DetermineReportingMonth(runDt As Date) As Date
    Dim yr As Long, mth As Long
    yr = Year(runDt)
    mth = Month(runDt)
    
    If Day(runDt) >= 10 Then
        ' On/after the 10th - process the previous month
        mth = mth - 1
    Else
        ' Before the 10th - process the month before previous
        mth = mth - 2
    End If
    
    ' Handle year boundary cases (December→January, February→January, etc.)
    ' DateSerial handles this automatically by wrapping correctly:
    ' - If month becomes 0 or negative, it will adjust to previous year
    ' - Example: DateSerial(2025, -2, 1) = DateSerial(2024, 10, 1) = October 1, 2024
    If mth <= 0 Then
        mth = mth + 12
        yr = yr - 1
    End If
    
    ' Return first day of the target month
    DetermineReportingMonth = DateSerial(yr, mth, 1)
End Function

' Function to check if a sheet with the given name exists
' Duplicated from mod_DataIO.SheetExists for encapsulation
Public Function SheetExists(shtName As String) As Boolean
    On Error Resume Next
    SheetExists = Len(ThisWorkbook.Sheets(shtName).Name) > 0
    On Error GoTo 0
End Function

' Function to refresh the 510k Power Query connection synchronously
Public Function Refresh510kConnection() As Boolean
    ' *** IMPORTANT: If you rename your Power Query in the QueryEditor, update this constant ***
    Const CN_NAME As String = "Query - pgGet510kData" ' Default connection name pattern
    Dim wsData As Worksheet
    Dim tblData As ListObject
    Dim cn As WorkbookConnection
    
    Refresh510kConnection = False ' Default return value
    
    On Error Resume Next
    
    ' Get the data worksheet
    Set wsData = ThisWorkbook.Sheets(mod_Config.DATA_SHEET_NAME)
    If wsData Is Nothing Then
        Debug.Print Time & " - ERROR: Could not find data sheet '" & mod_Config.DATA_SHEET_NAME & "'"
        Exit Function
    End If
    
    ' Get the table on that sheet
    Set tblData = wsData.ListObjects(1)
    If tblData Is Nothing Then
        Debug.Print Time & " - ERROR: Could not find data table on sheet '" & mod_Config.DATA_SHEET_NAME & "'"
        Exit Function
    End If
    
    ' Try the direct connection approach first
    Set cn = ThisWorkbook.Connections(CN_NAME)
    If Not cn Is Nothing Then
        ' Set to synchronous mode and refresh
        cn.OLEDBConnection.BackgroundQuery = False
        cn.Refresh
        Refresh510kConnection = True
    Else
        ' Fall back to mod_DataIO if the direct connection approach fails
        Refresh510kConnection = mod_DataIO.RefreshPowerQuery(tblData)
    End If
    
    On Error GoTo 0
End Function
