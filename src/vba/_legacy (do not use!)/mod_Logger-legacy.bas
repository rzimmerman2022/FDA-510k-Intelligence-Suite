'========================  mod_Logger  =====================================
Option Explicit

' --- Constants ---
Private Const LOG_SHEET As String = "RunLog"   ' Name of the hidden log sheet
Private Const BUF_CHUNK As Long = 256          ' Size to grow log buffer array
Private Const LOG_COL_COUNT As Long = 7        ' Number of columns in the log

' --- Public Enum for Log Levels ---
' Renamed from eLvl to eLogLevel to avoid ambiguity/conflicts with other possible declarations
Public Enum eLogLevel
    lvlINFO = 1    ' Standard information messages
    lvlDETAIL      ' Detailed debug information (conditional on DebugMode)
    lvlWARN        ' Warning conditions
    lvlERROR       ' Error conditions
End Enum

' --- Private State Variable (Module Level) ---
' Holds log data in memory before flushing
Private Type tBuf
    data()  As Variant   ' 2-D array (rows x LOG_COL_COUNT cols)
    used    As Long      ' Pointer to the next available row index
    runID   As String    ' Unique ID for the current execution session
End Type
Private logBuf As tBuf

' --- Public Subroutines ---

Public Sub LogEvt(stepTxt As String, level As eLogLevel, msg As String, _
           Optional extra As String = "")
    ' Purpose: Adds an event entry to the in-memory log buffer.
    ' Inputs:  stepTxt - Identifier for the process step (e.g., "ScoreLoop", "Archive").
    '          level - The severity level (lvlINFO, lvlDETAIL, lvlWARN, lvlERROR).
    '          msg - The main log message text.
    '          extra - Optional additional details.

    ' Optionally skip DETAIL level logging if DebugMode is not enabled
    If level = lvlDETAIL And Not DebugModeOn() Then Exit Sub

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
    ' Use direct array indexing rather than .item() method
    logBuf.data(logBuf.used, 1) = logBuf.runID         ' RunID
    logBuf.data(logBuf.used, 2) = Now                  ' Timestamp
    logBuf.data(logBuf.used, 3) = Environ$("USERNAME") ' User
    logBuf.data(logBuf.used, 4) = stepTxt              ' Step
    logBuf.data(logBuf.used, 5) = Choose(level, "INFO", "DETAIL", "WARN", "ERROR") ' Level Text
    logBuf.data(logBuf.used, 6) = msg                  ' Message
    logBuf.data(logBuf.used, 7) = extra                ' Extra Info
End Sub

Public Sub FlushLogBuf()
    ' Purpose: Writes the entire contents of the in-memory log buffer to the hidden RunLog sheet.
    '          Should be called once at the very end of the main process or before closing.

    ' Exit if buffer is empty (nothing to write)
    If logBuf.used = 0 Then Exit Sub

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo FlushError ' Handle errors during sheet writing

    ' Get the log sheet object (creates if it doesn't exist)
    Set ws = EnsureLogSheet
    If ws Is Nothing Then GoTo ResetBuffer ' Exit if sheet couldn't be created/found

    ' Find the next empty row on the log sheet (based on RunID column A)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Write the entire buffer to the sheet in one operation (fastest way)
    ' Using Application.Index to efficiently transfer used part of the buffer
    ws.Cells(nextRow, 1).Resize(logBuf.used, LOG_COL_COUNT).Value = _
        Application.Index(logBuf.data, Evaluate("ROW(1:" & logBuf.used & ")"), _
                           Evaluate("COLUMN(1:" & LOG_COL_COUNT & ")")) ' Using COLUMN avoids hardcoding 7

    Debug.Print Time & " - Flushed " & logBuf.used & " log entries to sheet '" & LOG_SHEET & "'."

ResetBuffer:
    ' Clear the buffer array and reset state variables after flushing
    Erase logBuf.data
    logBuf.used = 0
    logBuf.runID = "" ' Reset RunID for next session
    Exit Sub ' Normal exit

FlushError:
     Debug.Print Time & " - ERROR flushing log buffer to sheet: " & Err.Description
     ' Optionally attempt to reset buffer even on error
     Resume ResetBuffer
End Sub

Public Sub TrimRunLog(Optional keepRows As Long = 5000)
    ' Purpose: Deletes older rows from the log sheet to prevent unlimited growth.
    ' Inputs:  keepRows - The approximate number of most recent log entries to keep.

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
    ' We want to keep 'keepRows' number of data rows (plus header)
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

    ' Log the start of the session
    LogEvt "Logger", lvlINFO, "Session started.", "Version=" & mod_510k_Processor.VERSION_INFO ' Reference Public Const
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
            ' Set sheet to VeryHidden so user doesn't see it unless they know VBA
            ws.Visible = xlSheetVeryHidden
            ' Add header row
            ws.Range("A1").Resize(1, LOG_COL_COUNT).Value = Array("RunID", "Timestamp", "User", "Step", "Level", "Message", "Extra")
            ws.Range("A1").Resize(1, LOG_COL_COUNT).Font.Bold = True
            ' Optional: AutoFit columns after adding header
            ' ws.Columns("A:G").AutoFit
            Debug.Print Time & " - Created hidden log sheet: '" & LOG_SHEET & "'"
        Else
            ' Failed to create/rename sheet
            Debug.Print Time & " - ERROR: Could not create log sheet '" & LOG_SHEET & "'. Error: " & Err.Description
            Set ws = Nothing ' Return Nothing on failure
        End If
        On Error GoTo 0 ' Restore default error handling
    End If
    ' Return the worksheet object (either found or newly created)
    Set EnsureLogSheet = ws
End Function

Private Function DebugModeOn() As Boolean
    ' Purpose: Checks if detailed logging should be enabled.
    ' Checks: 1. Is current user the maintainer? AND 2. Is Named Range "DebugMode" set to TRUE?
    Dim debugName As Name
    Dim debugValue As String
    DebugModeOn = False ' Default to off

    ' Check if user is maintainer first (optimization)
    If Not mod_510k_Processor.IsMaintainerUser() Then Exit Function

    On Error Resume Next ' Handle error if named range doesn't exist
    Set debugName = ThisWorkbook.Names("DebugMode")
    If Err.Number = 0 Then
        ' Named range exists, get its value
        debugValue = LCase$(debugName.RefersToRange.Value2)
        If debugValue = "true" Then DebugModeOn = True
    Else
        ' Named range doesn't exist - log a warning once?
        Debug.Print Time & " - Warning: Named Range 'DebugMode' not found. DETAIL logging disabled."
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Function


