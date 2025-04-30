'========================  mod_Logger  =====================================
Option Explicit

' --- Constants ---
Private Const LOG_SHEET As String = "RunLog"   ' Name of the hidden log sheet
Private Const BUF_CHUNK As Long = 256          ' Size to grow log buffer array
Private Const LOG_COL_COUNT As Long = 7        ' Number of columns in the log

' --- Public Enum for Log Levels ---
' Using unique 'lg' prefix as requested to avoid conflicts
Public Enum LogLevel
    lgINFO = 1    ' Standard information messages
    lgDETAIL = 2  ' Detailed debug information (conditional on DebugMode)
    lgWARN = 3    ' Warning conditions
    lgERROR = 4   ' Error conditions
End Enum

' --- Private State Variable (Module Level) ---
' Holds log data in memory before flushing
Private Type tBuf
    data()  As Variant   ' 2-D array (rows x LOG_COL_COUNT cols)
    used    As Long      ' Pointer to the next available row index
    runID   As String    ' Unique ID for the current execution session
    isClosed As Boolean  ' Flag to prevent logging after fatal flush
End Type
Private logBuf As tBuf

' --- Public Subroutines ---

' Using the Public LogLevel Enum in the signature
Public Sub LogEvt(stepTxt As String, level As LogLevel, msg As String, _
           Optional extra As String = "")
    ' Purpose: Adds an event entry to the in-memory log buffer.
    ' Inputs:  stepTxt - Identifier for the process step.
    '          level - The severity level (using LogLevel Enum).
    '          msg - The main log message text.
    '          extra - Optional additional details.

    ' Prevent logging if logger is marked as closed
    If logBuf.isClosed Then Exit Sub

    ' Optionally skip DETAIL level logging if DebugMode is not enabled
    If level = LogLevel.lgDETAIL And Not DebugModeOn() Then Exit Sub ' Use LogLevel Enum

    ' Initialize logger (get RunID, create buffer) on first call per session
    If logBuf.runID = "" Then InitLogger

    ' Increment row pointer
    logBuf.used = logBuf.used + 1

    ' Check if buffer needs to be expanded
    If logBuf.used > UBound(logBuf.data, 1) Then
        ' Grow buffer by BUF_CHUNK rows, preserving existing data
        ReDim Preserve logBuf.data(1 To UBound(logBuf.data, 1) + BUF_CHUNK, 1 To LOG_COL_COUNT)
        Debug.Print Time & " - Log Buffer expanded." ' Optional debug message
    End If

    ' Write log data to the next available row in the buffer array
    logBuf.data(logBuf.used, 1) = logBuf.runID         ' RunID
    logBuf.data(logBuf.used, 2) = Now                  ' Timestamp
    logBuf.data(logBuf.used, 3) = Environ$("USERNAME") ' User
    logBuf.data(logBuf.used, 4) = stepTxt              ' Step
    ' Use Choose based on the Enum value
    logBuf.data(logBuf.used, 5) = Choose(level, "INFO", "DETAIL", "WARN", "ERROR") ' Level Text
    logBuf.data(logBuf.used, 6) = msg                  ' Message
    logBuf.data(logBuf.used, 7) = extra                ' Extra Info
End Sub

Public Sub FlushLogBuf()
    ' Purpose: Writes the entire contents of the in-memory log buffer to the hidden RunLog sheet.
    '          Should be called once at the very end of the main process or before closing.

    ' Don't try to flush if already closed
    If logBuf.isClosed Then Exit Sub

    ' Exit if buffer is empty (nothing to write)
    If logBuf.used = 0 Then Exit Sub

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo FlushError ' Handle errors during sheet writing

    ' Get the log sheet object (creates if it doesn't exist)
    Set ws = EnsureLogSheet
    If ws Is Nothing Then GoTo ResetBufferAndExit ' Exit if sheet couldn't be created/found

    ' Find the next empty row on the log sheet (based on RunID column A)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Write the entire buffer to the sheet in one operation (fastest way)
    Dim writeSuccess As Boolean: writeSuccess = False
    On Error Resume Next ' Handle potential #CALC! or other write errors
    ws.Cells(nextRow, 1).Resize(logBuf.used, LOG_COL_COUNT).value = _
        Application.Index(logBuf.data, Evaluate("ROW(1:" & logBuf.used & ")"), _
                           Evaluate("COLUMN(1:" & LOG_COL_COUNT & ")"))
    writeSuccess = (Err.Number = 0)
    On Error GoTo FlushError ' Restore primary error handler

    If writeSuccess Then
        Debug.Print Time & " - Flushed " & logBuf.used & " log entries to sheet '" & LOG_SHEET & "'."
    Else
        Debug.Print Time & " - FAILED to write log buffer to sheet '" & LOG_SHEET & "'. Error: " & Err.Description
        ' Attempt to write a simplified error message directly to the sheet
        On Error Resume Next
        ws.Cells(nextRow, 1).Resize(1, LOG_COL_COUNT).value = _
            Array(logBuf.runID, Now, Environ$("USERNAME"), "FlushError", "ERROR", "Failed to write log buffer via Application.Index", Err.Description)
        On Error GoTo 0
    End If

ResetBufferAndExit:
    ' Clear the buffer array and reset state variables after attempting flush
    Erase logBuf.data
    logBuf.used = 0
    logBuf.runID = "" ' Reset RunID for next session
    ' Leave logBuf.isClosed as False here for normal exit

    Exit Sub ' Normal exit

FlushError:
     Dim errDesc As String: errDesc = Err.Description ' Capture error info
     Debug.Print Time & " - ERROR during FlushLogBuf processing: " & errDesc
     ' Set isClosed flag on fatal error during flush
     logBuf.isClosed = True
     Debug.Print Time & " - Logger marked as closed due to fatal flush error."
     ' Optionally attempt to reset buffer even on error
     Resume ResetBufferAndExit
End Sub

Public Sub TrimRunLog(Optional keepRows As Long = 5000)
    ' Purpose: Deletes older rows from the log sheet to prevent unlimited growth.
    ' Inputs:  keepRows - The approximate number of most recent log entries to keep.

    ' Don't try to trim if logger closed
    If logBuf.isClosed Then
        Debug.Print Time & " - Log Trim skipped as logger is marked closed."
        Exit Sub
    End If

    Const LOG_SHEET As String = "RunLog" ' Ensure constant is defined or accessible
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstRowToDelete As Long

    On Error GoTo TrimError

    ' Get the log sheet object
    Set ws = EnsureLogSheet
    If ws Is Nothing Then Exit Sub ' Cannot trim if sheet doesn't exist

    ' Find the last row with data in Column A (RunID)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Calculate the first row number to delete (keeping header row 1)
    firstRowToDelete = lastRow - keepRows + 1

    ' Only delete if we have more rows than we want to keep (and more than just the header)
    If firstRowToDelete > 2 Then
        Application.ScreenUpdating = False ' Prevent screen flicker during delete
        ws.Rows("2:" & firstRowToDelete - 1).Delete
        Application.ScreenUpdating = True
        Debug.Print Time & " - Trimmed log sheet. Deleted rows 2 through " & firstRowToDelete - 1 & "."
    Else
        Debug.Print Time & " - Log sheet trim skipped (Rows <= KeepRows)."
    End If
    Exit Sub

TrimError:
     Application.ScreenUpdating = True ' Ensure screen updating is re-enabled
     Debug.Print Time & " - Error trimming log sheet: " & Err.Description
End Sub


'------------------------- Internal Helper Functions -------------------------

Private Sub InitLogger()
    ' Purpose: Initializes the log buffer for a new session. Gets RunID, allocates initial array size.
    On Error Resume Next ' In case CreateObject fails
    logBuf.runID = CreateObject("Scriptlet.TypeLib").GUID ' Generate unique session ID
    If Err.Number <> 0 Then logBuf.runID = "ErrorGUID_" & Format(Now, "yyyymmddhhmmss") ' Fallback ID
    Err.Clear
    On Error GoTo 0

    ' Pre-allocate the first chunk of the buffer array
    ReDim logBuf.data(1 To BUF_CHUNK, 1 To LOG_COL_COUNT)
    logBuf.used = 0 ' Reset row pointer
    logBuf.isClosed = False ' Ensure logger is open on init

    ' Log the start of the session using the new Enum member
    ' Note: This call itself uses LogEvt, so it must be correct.
    ' Assuming mod_510k_Processor.VERSION_INFO is accessible
    On Error Resume Next ' Avoid error if VERSION_INFO isn't ready yet
    LogEvt "Logger", LogLevel.lgINFO, "Session started.", "Version=" & mod_510k_Processor.VERSION_INFO
    On Error GoTo 0
End Sub

Private Function EnsureLogSheet() As Worksheet
    ' Purpose: Finds the log sheet or creates it if it doesn't exist.
    ' Returns: Worksheet object for the log sheet, or Nothing on failure.
    Dim ws As Worksheet
    On Error Resume Next ' Temporarily ignore error if sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0 ' Restore default error handling

    If ws Is Nothing Then
        ' Sheet doesn't exist, try to create it
        On Error Resume Next ' Handle errors during sheet creation/naming
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Err.Number = 0 Then
            ws.Name = LOG_SHEET
            ws.Visible = xlSheetVeryHidden
            ws.Range("A1").Resize(1, LOG_COL_COUNT).value = Array("RunID", "Timestamp", "User", "Step", "Level", "Message", "Extra")
            ws.Range("A1").Resize(1, LOG_COL_COUNT).Font.Bold = True
            Debug.Print Time & " - Created hidden log sheet: '" & LOG_SHEET & "'"
        Else
            Debug.Print Time & " - ERROR: Could not create log sheet '" & LOG_SHEET & "'. Error: " & Err.Description
            Set ws = Nothing ' Return Nothing on failure
        End If
        On Error GoTo 0 ' Restore default error handling
    End If
    Set EnsureLogSheet = ws
End Function

Private Function DebugModeOn() As Boolean
    ' Purpose: Checks if detailed logging (lgDETAIL) should be enabled.
    Dim debugName As Name
    Dim debugValue As String
    DebugModeOn = False ' Default to off

    ' Check Environment Variable Override First
    On Error Resume Next
    If Environ$("TRACE_ALL_USERS") = "1" Then
        DebugModeOn = True
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0 ' Restore default error handling if needed

    ' Check Maintainer Status
    If Not mod_510k_Processor.IsMaintainerUser() Then Exit Function

    ' Check Named Range "DebugMode"
    On Error Resume Next ' Handle error if named range doesn't exist
    Set debugName = ThisWorkbook.Names("DebugMode")
    If Err.Number = 0 Then
        ' Named range exists, get its value
        debugValue = LCase$(Trim(CStr(debugName.RefersToRange.Value2)))
        If debugValue = "true" Then DebugModeOn = True
    Else
        ' Named range doesn't exist - log a warning? (Maybe not needed if Maintainer check passed)
        Debug.Print Time & " - Warning: Named Range 'DebugMode' not found. DETAIL logging relies on Maintainer status only."
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Function
