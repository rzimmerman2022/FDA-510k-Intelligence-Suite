' ==========================================================================
' Module      : mod_Utils
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module provides a collection of miscellaneous utility
'               functions that are used across various modules within the
'               510(k) Intelligence Suite application. These functions handle
'               common tasks such as safely retrieving worksheet objects,
'               checking user permissions (maintainer status), ensuring Excel
'               UI settings are restored correctly, and performing calculations
'               like determining color brightness for contrast purposes.
'
' Key Functions:
'               - GetWorksheets: Safely retrieves Worksheet objects for the
'                 main data, weights, and cache sheets, handling errors if
'                 any are missing.
'               - IsMaintainerUser: Checks if the current Windows username
'                 matches the MAINTAINER_USERNAME constant defined in mod_Config.
'               - EnsureUIOn: Restores Application settings (ScreenUpdating,
'                 StatusBar, Cursor, Calculation, EnableEvents) typically
'                 called at the end of processes or in error handlers.
'               - GetBrightness: Calculates the perceived brightness of an
'                 RGB color value, useful for determining appropriate font color
'                 (e.g., black or white) for readability against a background.
'
' Dependencies: - mod_Logger: For logging errors (e.g., missing worksheets).
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - mod_Config: Relies on sheet name constants (DATA_SHEET_NAME, etc.)
'                 and MAINTAINER_USERNAME.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Corrected module reference for IsMaintainerUser
'                               call within mod_Logger.DebugModeOn.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit
Attribute VB_Name = "mod_Utils"

Public Function GetWorksheets(ByRef wsData As Worksheet, ByRef wsWeights As Worksheet, ByRef wsCache As Worksheet) As Boolean
    ' Purpose: Safely gets required worksheet objects, logs errors if not found.
    Const PROC_NAME As String = "mod_Utils.GetWorksheets" ' Updated PROC_NAME
    GetWorksheets = False ' Default to failure
    Dim missingSheets As String

    On Error Resume Next ' Check each sheet individually
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    If Err.Number <> 0 Then missingSheets = missingSheets & vbCrLf & " - " & DATA_SHEET_NAME: Err.Clear
    Set wsWeights = ThisWorkbook.Sheets(WEIGHTS_SHEET_NAME)
    If Err.Number <> 0 Then missingSheets = missingSheets & vbCrLf & " - " & WEIGHTS_SHEET_NAME: Err.Clear
    Set wsCache = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    If Err.Number <> 0 Then missingSheets = missingSheets & vbCrLf & " - " & CACHE_SHEET_NAME: Err.Clear
    On Error GoTo 0 ' Restore default error handling

    If Len(missingSheets) > 0 Then
        LogEvt PROC_NAME, lgERROR, "Required worksheet(s) not found:" & Replace(missingSheets, vbCrLf, ", ")
        TraceEvt lvlERROR, PROC_NAME, "Required worksheet(s) not found", "Missing=" & Replace(missingSheets, vbCrLf, ", ")
        MsgBox "Error: The following required worksheets could not be found:" & missingSheets & vbCrLf & "Please ensure the sheets exist and names match the configuration.", vbCritical, "Missing Worksheets"
        ' Call EnsureUIOn here to prevent leaving the UI in a bad state
        Call EnsureUIOn(xlCalculationAutomatic) ' Restore to automatic calc on critical error
    Else
        LogEvt PROC_NAME, lgDETAIL, "All required worksheets found."
        TraceEvt lvlDET, PROC_NAME, "All required worksheets found"
        GetWorksheets = True ' Success
    End If
End Function

Public Function IsMaintainerUser() As Boolean
    ' Purpose: Checks if the current user matches the configured maintainer username.
    '          Used to enable/disable features like OpenAI calls or bypassing archive checks.
    IsMaintainerUser = (LCase(Environ("USERNAME")) = LCase(MAINTAINER_USERNAME))
End Function

Public Sub EnsureUIOn(Optional restoreCalcState As XlCalculation = xlCalculationManual)
    ' Purpose: Restores standard Excel UI settings after processing or on error.
    '          Should be called in error handlers and at the end of main routines.
    Const PROC_NAME As String = "mod_Utils.EnsureUIOn" ' Updated PROC_NAME
    On Error Resume Next ' Prevent errors within this cleanup routine from stopping everything
    Application.ScreenUpdating = True
    Application.StatusBar = False ' Clear status bar
    Application.Cursor = xlDefault
    Application.Calculation = restoreCalcState ' Restore original or specified calculation state
    Application.EnableEvents = True
    TraceEvt lvlDET, PROC_NAME, "UI settings restored", "CalcState=" & restoreCalcState
    On Error GoTo 0
End Sub

Public Function GetBrightness(rgbColor As Long) As Double
    ' Purpose: Calculates the perceived brightness of an RGB color. Used for text contrast.
    ' Formula: (0.299*R + 0.587*G + 0.114*B) / 255
    Dim R As Long, G As Long, B As Long
    R = rgbColor Mod 256
    G = (rgbColor \ 256) Mod 256
    B = (rgbColor \ 65536) Mod 256
    GetBrightness = (0.299 * R + 0.587 * G + 0.114 * B) / 255
End Function
