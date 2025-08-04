'=====  mod_DebugTraceHelpers  ===========================================
Option Explicit

'--  MASTER SWITCH --------------------------------------------------------
' *** Set TRACE_ENABLED to True for detailed logging, False for production ***
Public Const TRACE_ENABLED As Boolean = True    '<<< MASTER DEBUG SWITCH
' *** Set TRACE_LEVEL to control verbosity (higher number = more detail) ***
Public Const TRACE_LEVEL  As Long = 4           '<<< CHANGED: 3 to 4 for stabilisation (per review)
                                                '0=Off, 1=Error, 2=Warn, 3=Info, 4=Detail, 5=Spam/Loop

'--  target worksheet (created on-the-fly) --------------------------------
Private Const TRACE_SHEET As String = "DebugTrace"

'--  public enum for readability -----------------------------------------
Public Enum eTraceLvl '<<< Keeping this Public as it resolved the ambiguity
    lvlOFF = 0   ' Added for completeness
    lvlERROR = 1
    lvlWARN = 2
    lvlINFO = 3
    lvlDET = 4   ' Detail
    lvlSPAM = 5  ' Loop/Spammy
End Enum

'--------------------------------------------------------------------------
' Procedure : TraceEvt
' Author    : (Adapted from prompt)
' Date      : 2025-04-29
' Purpose   : Writes a log entry to the TRACE_SHEET if enabled and level permits.
' Parameters: lvl - The severity level of the event (eTraceLvl enum).
'             proc - The name of the procedure where the event occurred.
'             msg - The main message describing the event.
'             detail - Optional additional details about the event.
'--------------------------------------------------------------------------
Public Sub TraceEvt(ByVal lvl As eTraceLvl, _
                    ByVal proc As String, _
                    ByVal msg As String, _
           Optional ByVal detail As String = vbNullString)

    If Not TRACE_ENABLED Then Exit Sub
    If lvl = lvlOFF Or lvl > TRACE_LEVEL Then Exit Sub ' Check against master level

    Dim ws As Worksheet
    Dim wsExists As Boolean
    Dim r As Range

    ' --- Attempt to get the sheet ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRACE_SHEET)
    wsExists = (Err.Number = 0)
    On Error GoTo 0 ' Restore error handling

    ' --- Create sheet if it doesn't exist ---
    If Not wsExists Then
        On Error GoTo TraceErrorHandler ' Handle errors during sheet creation/setup
        Application.ScreenUpdating = False ' Prevent flicker if creating sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = TRACE_SHEET
        ' Set up headers
        With ws.Range("A1:E1") ' Added Column E for Detail
            .value = Array("Timestamp", "Level", "Procedure", "Message", "Details")
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220) ' Light grey header
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        ws.Columns("A:E").AutoFit ' Autofit initially
        ws.Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss.000" ' Precise timestamp
        ws.Columns("D").ColumnWidth = 60 ' Give message more space
        ws.Columns("E").ColumnWidth = 50 ' Give details space
        Application.ScreenUpdating = True
        On Error GoTo 0 ' Restore error handling after successful creation
    End If

    ' --- Write the log entry ---
    On Error GoTo TraceErrorHandler ' Handle errors during writing
    With ws
        ' Find the next empty row in Column A
        Set r = .Cells(.Rows.Count, "A").End(xlUp).offset(1)
        ' Write data using Resize
        r.Resize(1, 5).value = Array(Now, LevelName(lvl), proc, msg, detail)
        ' Optional: Apply light banding for readability
        If r.Row Mod 2 = 0 Then r.Resize(1, 5).Interior.Color = RGB(245, 245, 245)

        ' <<< ADDED: Autofit Message/Detail columns periodically to handle long text >>>
        If r.Row Mod 100 = 0 Then
             On Error Resume Next ' Don't let autofit error stop logging
             ws.Columns("D:E").AutoFit
             On Error GoTo TraceErrorHandler ' Restore handler
        End If
        ' <<< END ADDITION >>>

    End With

    Set ws = Nothing
    Set r = Nothing
    Exit Sub

TraceErrorHandler:
    ' Basic error handler to prevent log failures from crashing the main code
    Debug.Print Now & " - ERROR in TraceEvt (mod_DebugTraceHelpers): " & Err.Number & " - " & Err.Description
    ' Optionally, try to write a simplified error to the sheet if possible
    On Error Resume Next
    If Not ws Is Nothing Then
         Set r = ws.Cells(ws.Rows.Count, "A").End(xlUp).offset(1)
         r.Resize(1, 5).value = Array(Now, "LOG_ERR", "TraceEvt", "Error writing log entry", Err.Description)
         r.Resize(1, 5).Font.Color = vbRed
    End If
    On Error GoTo 0 ' Prevent infinite loop if error handler itself fails
    Set ws = Nothing
    Set r = Nothing
End Sub

'--------------------------------------------------------------------------
' Function  : LevelName
' Author    : (Adapted from prompt)
' Date      : 2025-04-29
' Purpose   : Converts the eTraceLvl enum to a short string representation.
'--------------------------------------------------------------------------
Private Function LevelName(lvl As eTraceLvl) As String
    Select Case lvl
        Case lvlERROR: LevelName = "ERROR"
        Case lvlWARN:  LevelName = "WARN"
        Case lvlINFO:  LevelName = "INFO"
        Case lvlDET:   LevelName = "DETAIL"
        Case lvlSPAM:  LevelName = "SPAM"
        Case Else:     LevelName = "LVL_" & CStr(lvl) ' Handle unexpected values
    End Select
End Function

'--------------------------------------------------------------------------
' Procedure : ClearDebugTrace
' Author    : (From prompt)
' Date      : 2025-04-29
' Purpose   : Clears all log entries from the DebugTrace sheet, leaving headers.
'--------------------------------------------------------------------------
Sub ClearDebugTrace()
    Dim ws As Worksheet
    On Error Resume Next ' Ignore error if sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets(TRACE_SHEET)
    If Err.Number = 0 Then
        With ws
            If .FilterMode Then .ShowAllData ' Clear filters if active
            .Range("A2:" & .Cells(.Rows.Count, .Columns.Count).Address).ClearContents
            .Range("A2").Select ' Optional: Select top cell after clearing
        End With
        MsgBox "'" & TRACE_SHEET & "' cleared.", vbInformation
    Else
        MsgBox "'" & TRACE_SHEET & "' sheet not found.", vbExclamation
    End If
    On Error GoTo 0
    Set ws = Nothing
End Sub

'==========================================================================
