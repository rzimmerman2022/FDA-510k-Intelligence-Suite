' ==========================================================================
' Module      : mod_Logger
' Author      : Unknown / Updated by AI Assistant for User
' Date        : 2025-05-07 (Update)
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO (Logger internal version 2.0)
' ==========================================================================
' Description : Provides a buffered logging system to write diagnostic and
'               event information to a hidden Excel sheet ('RunLog').
'               Includes different log levels and automatic session tracking.
'
' Key Functions/Procedures:
'               - LogEvt: Adds a log entry to the in-memory buffer.
'               - FlushLogBuf: Writes the entire buffer to the 'RunLog' sheet.
'               - TrimRunLog: Deletes older rows from the 'RunLog' sheet.
'
' Updates in this version:
'               - Added lgFATAL to LogLevel enum.
'               - Updated LogEvt to handle lgFATAL string representation.
'               - Added conditional compilation for internal Debug.Print statements.
' ==========================================================================
Option Explicit

' --- Conditional Compilation for Debugging this Module ---
#Const DEBUG_LOGGER = False ' Set to True to enable Debug.Print statements within this module

' --- Constants ---
Private Const LOG_SHEET_NAME As String = "RunLog" ' Name of the hidden log sheet
Private Const BUF_CHUNK As Long = 256           ' Size to grow log buffer array
Private Const LOG_COL_COUNT As Long = 7         ' Number of columns in the log

' --- Public Enum for Log Levels ---
' Using unique 'lg' prefix as requested to avoid conflicts
Public Enum LogLevel
    lgINFO = 1      ' Standard information messages
    lgDETAIL = 2    ' Detailed debug information (conditional on DebugModeOn function)
    lgWARN = 3      ' Warning conditions
    lgERROR = 4     ' Error conditions
    lgFATAL = 5     ' Fatal error conditions (NEW)
    ' Add lgDEBUG = 0 if needed, but Choose function would need adjustment
End Enum

' --- Private State Variable (Module Level) ---
' Holds log data in memory before flushing
Private Type tBuf
    data()    As Variant  ' 2-D array (rows x LOG_COL_COUNT cols)
    used      As Long     ' Pointer to the next available row index
    runID     As String   ' Unique ID for the current execution session
    isClosed  As Boolean  ' Flag to prevent logging after fatal flush
    isInit    As Boolean  ' Flag to ensure InitLogger runs only once effectively per session
End Type
Private logBuf As tBuf

' --- Public Subroutines ---

Public Sub LogEvt(stepTxt As String, ByVal level As LogLevel, msg As String, _
                  Optional extra As String = "")
    ' Purpose: Adds an event entry to the in-memory log buffer.
    ' Inputs:  stepTxt - Identifier for the process step/event name.
    '          level - The severity level (using LogLevel Enum).
    '          msg - The main log message text.
    '          extra - Optional additional details.

#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.LogEvt: Entered. Step=" & stepTxt & ", Level=" & CStr(level) & ", Msg=" & Left(msg, 100)
#End If

    ' Prevent logging if logger is marked as closed
    If logBuf.isClosed Then Exit Sub

    ' Optionally skip DETAIL level logging if DebugMode is not enabled
    If level = LogLevel.lgDETAIL And Not DebugModeOn() Then Exit Sub

    ' Initialize logger (get RunID, create buffer) on first call per session
    If Not logBuf.isInit Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.LogEvt: Calling InitLogger..."
#End If
        InitLogger
        ' Check if InitLogger failed silently (e.g., runID still empty or logger closed)
        If Not logBuf.isInit Or logBuf.isClosed Then
#If DEBUG_LOGGER Then
            Debug.Print Now & " - mod_Logger.LogEvt: Exiting because InitLogger failed or logger is closed post-init."
#End If
            Exit Sub
        End If
    End If

    ' Increment row pointer
    logBuf.used = logBuf.used + 1

    ' Check if buffer needs to be expanded
    If logBuf.used > UBound(logBuf.data, 1) Then
        ' Grow buffer by BUF_CHUNK rows, preserving existing data
        ReDim Preserve logBuf.data(1 To UBound(logBuf.data, 1) + BUF_CHUNK, 1 To LOG_COL_COUNT)
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.LogEvt: Log buffer resized to " & UBound(logBuf.data, 1) & " rows."
#End If
    End If

    ' Write log data to the next available row in the buffer array
    logBuf.data(logBuf.used, 1) = logBuf.runID         ' RunID
    logBuf.data(logBuf.used, 2) = Now                  ' Timestamp
    logBuf.data(logBuf.used, 3) = Environ$("USERNAME") ' User
    logBuf.data(logBuf.used, 4) = stepTxt              ' Step/Event Name
    
    ' Get Level Text - ensure Choose covers all defined LogLevel enum members starting from 1
    On Error Resume Next ' Handle if level is out of Choose range temporarily
    Dim levelString As String
    levelString = Choose(level, "INFO", "DETAIL", "WARN", "ERROR", "FATAL") ' Added FATAL
    If Err.Number <> 0 Then
        levelString = "LEVEL_" & CStr(level) ' Fallback if level is somehow out of expected range
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling for the sub

    logBuf.data(logBuf.used, 5) = levelString          ' Level Text
    logBuf.data(logBuf.used, 6) = msg                  ' Message
    logBuf.data(logBuf.used, 7) = extra                ' Extra Info
End Sub

Public Sub FlushLogBuf()
    ' Purpose: Writes the entire contents of the in-memory log buffer to the hidden RunLog sheet.
    '          Should be called once at the very end of the main process or before closing.
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.FlushLogBuf: Entered. isClosed=" & logBuf.isClosed & ", UsedRows=" & logBuf.used
#End If

    ' Don't try to flush if already closed or buffer empty
    If logBuf.isClosed Or logBuf.used = 0 Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.FlushLogBuf: Exiting because logger closed or buffer empty."
#End If
        Exit Sub
    End If

    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim originalCalc As XlCalculation

    On Error GoTo FlushError_Handler ' Handle errors during sheet writing

    originalCalc = Application.Calculation
    Application.Calculation = xlCalculationManual ' For performance during write

#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.FlushLogBuf: Calling EnsureLogSheet..."
#End If
    Set wsLog = EnsureLogSheet
    If wsLog Is Nothing Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.FlushLogBuf: EnsureLogSheet returned Nothing. Cannot flush."
#End If
        logBuf.isClosed = True ' Prevent further logging attempts if sheet is unavailable
        GoTo ResetBufferAndExit_Flush ' Still reset buffer
    End If

    ' Find the next empty row on the log sheet (based on RunID column A)
    If wsLog.Cells(wsLog.Rows.Count, 1).value = "" Then ' If last cell in col A is blank
        nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
        If wsLog.Cells(nextRow, 1).value <> "" Then nextRow = nextRow + 1 ' If End(xlUp) found content, next row
        If nextRow = 1 And wsLog.Cells(1, 1).value = "" Then nextRow = 1 ' True first row
        If nextRow = 2 And wsLog.Cells(1, 1).value <> "" And wsLog.Cells(2, 1).value = "" Then nextRow = 2 ' Header exists, write to row 2

    Else ' Last cell in Col A has content, so next row is fine
        nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ' Ensure nextRow is at least 2 if header exists, or 1 if sheet is completely blank
    If nextRow = 1 And wsLog.Range("A1").value <> "" Then nextRow = 2


    ' Write the entire buffer to the sheet in one operation
    Dim writeSuccess As Boolean: writeSuccess = False
    Dim arrToWrite As Variant
    
    ' Create an array slice to write only used rows
    ReDim arrToWrite(1 To logBuf.used, 1 To LOG_COL_COUNT)
    Dim r As Long, c As Long
    For r = 1 To logBuf.used
        For c = 1 To LOG_COL_COUNT
            arrToWrite(r, c) = logBuf.data(r, c)
        Next c
    Next r

    On Error Resume Next ' Handle potential write errors specifically for this block
    wsLog.Cells(nextRow, 1).Resize(logBuf.used, LOG_COL_COUNT).value = arrToWrite
    writeSuccess = (Err.Number = 0)
    Dim writeErrNum As Long: writeErrNum = Err.Number
    Dim writeErrDesc As String: writeErrDesc = Err.Description
    On Error GoTo FlushError_Handler ' Restore primary error handler

    If Not writeSuccess Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.FlushLogBuf: Failed to write log buffer. Err=" & writeErrNum & ": " & writeErrDesc
#End If
        ' Attempt to write a simplified error message directly to the sheet
        On Error Resume Next
        wsLog.Cells(nextRow, 1).Resize(1, LOG_COL_COUNT).value = _
            Array(logBuf.runID, Now, Environ$("USERNAME"), "FlushLogBuf_WriteError", "ERROR", "Failed to write full log buffer. Err: " & writeErrNum, Left(writeErrDesc, 250))
        On Error GoTo 0
    Else
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.FlushLogBuf: Successfully wrote " & logBuf.used & " rows to " & wsLog.Name & " starting at row " & nextRow
#End If
    End If

ResetBufferAndExit_Flush:
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.FlushLogBuf: Reached ResetBufferAndExit_Flush."
#End If
    ' Clear the buffer array and reset state variables after attempting flush
    Erase logBuf.data
    logBuf.used = 0
    ' logBuf.runID = "" ' Keep RunID if session might continue, or clear if this is end of app
    logBuf.isInit = False ' Allow re-initialization if app continues and calls LogEvt again
    ' Leave logBuf.isClosed as is (might have been set true on critical error)
    Application.Calculation = originalCalc
    Exit Sub

FlushError_Handler:
    Dim errNumLocal As Long: errNumLocal = Err.Number
    Dim errDescLocal As String: errDescLocal = Err.Description
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.FlushLogBuf: FATAL ERROR HANDLER. Err=" & errNumLocal & ", Desc=" & errDescLocal
#End If
    logBuf.isClosed = True ' Set isClosed flag on fatal error during flush
    Application.Calculation = originalCalc ' Restore calculation
    Resume ResetBufferAndExit_Flush ' Attempt to reset buffer
End Sub

Public Sub TrimRunLog(Optional ByVal keepRows As Long = 5000)
    ' Purpose: Deletes older rows from the log sheet to prevent unlimited growth.
    ' Inputs:  keepRows - The approximate number of most recent log entries to keep.

    If logBuf.isClosed Then Exit Sub ' Don't try to trim if logger closed

    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim rowsToDelete As Long
    Dim originalScreenUpdating As Boolean

    On Error GoTo TrimError_Handler

    Set wsLog = EnsureLogSheet
    If wsLog Is Nothing Then Exit Sub ' Cannot trim if sheet doesn't exist

    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

    ' We keep header (row 1) + keepRows. So delete if lastRow > keepRows + 1
    If lastRow > (keepRows + 1) Then
        rowsToDelete = lastRow - keepRows - 1 ' Number of data rows to delete from the top
        If rowsToDelete > 0 Then
            originalScreenUpdating = Application.ScreenUpdating
            Application.ScreenUpdating = False
            wsLog.Rows("2:" & rowsToDelete + 1).Delete ' Delete rows starting from row 2
            Application.ScreenUpdating = originalScreenUpdating
#If DEBUG_LOGGER Then
            Debug.Print Now & " - mod_Logger.TrimRunLog: Deleted " & rowsToDelete & " old log rows from '" & wsLog.Name & "'."
#End If
        End If
    End If
    Exit Sub

TrimError_Handler:
    If originalScreenUpdating Then Application.ScreenUpdating = True ' Ensure re-enabled
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.TrimRunLog: Error during trim. Err=" & Err.Number & ", Desc=" & Err.Description
#End If
    ' Log error? Usually Trim is non-critical.
End Sub


'------------------------- Internal Helper Functions -------------------------

Private Sub InitLogger()
    ' Purpose: Initializes the log buffer for a new session.
    On Error GoTo InitError_Handler
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.InitLogger: Entered."
#End If

    If logBuf.isInit And Not logBuf.isClosed Then Exit Sub ' Already initialized and not closed

    Dim guidErrNum As Long
    Dim guidErrDesc As String

    On Error Resume Next ' In case CreateObject fails
    logBuf.runID = CreateObject("Scriptlet.TypeLib").GUID ' Generate unique session ID
    guidErrNum = Err.Number
    guidErrDesc = Err.Description
    On Error GoTo InitError_Handler ' Restore main error handler for this sub

    If guidErrNum <> 0 Then
        logBuf.runID = "ErrorGUID_" & Format(Now, "yyyymmddhhmmss") ' Fallback ID
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.InitLogger: Failed to create GUID. Using fallback ID. Err=" & guidErrNum & ", Desc=" & guidErrDesc
#End If
    Else
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.InitLogger: Successfully created GUID: " & logBuf.runID
#End If
    End If

    ReDim logBuf.data(1 To BUF_CHUNK, 1 To LOG_COL_COUNT)
    logBuf.used = 0
    logBuf.isClosed = False
    logBuf.isInit = True ' Mark as initialized

    ' Log the start of the session
    ' This LogEvt call will not re-call InitLogger because logBuf.isInit is now True
    On Error Resume Next ' Avoid error if VERSION_INFO isn't ready yet or LogEvt fails early in a weird state
    Dim versionInfoStr As String
    versionInfoStr = "N/A"
    
    ' Try to access mod_Config.VERSION_INFO directly, but handle potential error if module is missing
    ' or VERSION_INFO isn't defined yet
    On Error Resume Next
    versionInfoStr = mod_Config.VERSION_INFO ' Assumes mod_Config.VERSION_INFO is accessible
    If Err.Number <> 0 Then versionInfoStr = "mod_Config.VERSION_INFO Error"
    On Error GoTo InitError_Handler
    
    LogEvt "LoggerInit", LogLevel.lgINFO, "Logger session initialized.", "RunID=" & logBuf.runID & ", Version=" & versionInfoStr
    If Err.Number <> 0 Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.InitLogger: Error during initial LogEvt call. Err=" & Err.Number & ", Desc=" & Err.Description
#End If
    End If
    On Error GoTo InitError_Handler ' Restore main error handler
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.InitLogger: Exiting normally."
#End If
    Exit Sub

InitError_Handler:
    Dim errNumLocal As Long: errNumLocal = Err.Number
    Dim errDescLocal As String: errDescLocal = Err.Description
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.InitLogger: FATAL ERROR HANDLER. Err=" & errNumLocal & ", Desc=" & errDescLocal
#End If
    If logBuf.runID = "" Then logBuf.runID = "INIT_LOGGER_FAILED_" & Format(Now, "yyyymmddhhmmss")
    logBuf.isClosed = True
    logBuf.isInit = True ' Mark as init attempted, even if failed, to stop recursion
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.InitLogger: Marked logger as closed and init attempted due to error."
#End If
End Sub

Private Function EnsureLogSheet() As Worksheet
    ' Purpose: Finds the log sheet or creates it if it doesn't exist.
    ' Returns: Worksheet object for the log sheet, or Nothing on failure.
#If DEBUG_LOGGER Then
    Debug.Print Now & " - mod_Logger.EnsureLogSheet: Entered. Looking for sheet '" & LOG_SHEET_NAME & "'."
#End If
    Dim ws As Worksheet
    Dim findErrNum As Long
    Dim findErrDesc As String

    On Error Resume Next ' Temporarily ignore error if sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    findErrNum = Err.Number
    findErrDesc = Err.Description
    On Error GoTo 0 ' Restore default error handling

    If ws Is Nothing Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.EnsureLogSheet: Sheet '" & LOG_SHEET_NAME & "' not found (Err=" & findErrNum & "). Attempting creation."
#End If
        Dim createErrNum As Long
        Dim createErrDesc As String
        Dim tempSheet As Worksheet
        
        On Error Resume Next ' Handle errors during sheet creation/naming
        Set tempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        createErrNum = Err.Number
        createErrDesc = Err.Description

        If createErrNum = 0 And Not tempSheet Is Nothing Then
            tempSheet.Name = LOG_SHEET_NAME
            createErrNum = Err.Number ' Capture error from rename attempt
            createErrDesc = Err.Description
            If createErrNum = 0 Then
                tempSheet.Visible = xlSheetVeryHidden
                tempSheet.Range("A1").Resize(1, LOG_COL_COUNT).value = Array("RunID", "Timestamp", "User", "Step", "Level", "Message", "Extra")
                tempSheet.Range("A1").Resize(1, LOG_COL_COUNT).Font.Bold = True
                Set ws = tempSheet ' Assign to ws only on full success
#If DEBUG_LOGGER Then
                Debug.Print Now & " - mod_Logger.EnsureLogSheet: Successfully created and setup sheet '" & LOG_SHEET_NAME & "'."
#End If
            Else
#If DEBUG_LOGGER Then
                Debug.Print Now & " - mod_Logger.EnsureLogSheet: FAILED to rename new sheet to '" & LOG_SHEET_NAME & "'. Err=" & createErrNum & ", Desc=" & createErrDesc
#End If
                ' Attempt to delete the partially created sheet if rename failed
                Application.DisplayAlerts = False
                tempSheet.Delete
                Application.DisplayAlerts = True
                Set ws = Nothing ' Ensure Nothing is returned
            End If
        Else
#If DEBUG_LOGGER Then
            Debug.Print Now & " - mod_Logger.EnsureLogSheet: FAILED to add new sheet. Err=" & createErrNum & ", Desc=" & createErrDesc
#End If
            Set ws = Nothing ' Ensure Nothing is returned
        End If
        On Error GoTo 0 ' Restore default error handling
    Else
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.EnsureLogSheet: Found existing sheet '" & LOG_SHEET_NAME & "'."
#End If
    End If
    Set EnsureLogSheet = ws
End Function

Private Function DebugModeOn() As Boolean
    ' Purpose: Checks if detailed logging (lgDETAIL) should be enabled.
    ' Priority: Environment Variable > Maintainer User + Named Range "DebugMode"
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
    On Error GoTo 0

    ' Check Maintainer Status (Requires mod_Utils.IsMaintainerUser to exist)
    Dim isMaintainer As Boolean
    On Error Resume Next ' In case mod_Utils or IsMaintainerUser is missing
    isMaintainer = mod_Utils.IsMaintainerUser()
    If Err.Number <> 0 Then
#If DEBUG_LOGGER Then
        Debug.Print Now & " - mod_Logger.DebugModeOn: Error calling mod_Utils.IsMaintainerUser. Err=" & Err.Number & ", Desc=" & Err.Description
#End If
        isMaintainer = False ' Assume not maintainer if util fails
        Err.Clear
    End If
    On Error GoTo 0
    
    If Not isMaintainer Then Exit Function ' Not a maintainer, so debug mode via named range is off

    ' Maintainer: Check Named Range "DebugMode"
    On Error Resume Next
    Set debugName = ThisWorkbook.Names("DebugMode")
    If Err.Number = 0 Then
        debugValue = LCase$(Trim(CStr(debugName.RefersToRange.Value2)))
        If debugValue = "true" Or debugValue = "1" Or debugValue = "yes" Then
            DebugModeOn = True
        End If
    Else
        Err.Clear ' Named range "DebugMode" likely doesn't exist, which is fine.
    End If
    On Error GoTo 0
End Function
