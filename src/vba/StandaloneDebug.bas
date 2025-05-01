'==========================================================================
' Module      : StandaloneDebug
' Author      : Claude
' Date        : 2025-04-30
' Description : Simple standalone debugging module with no dependencies
'==========================================================================
Option Explicit
Attribute VB_Name = "StandaloneDebug"

' Create a completely new debug sheet with a timestamp
Private Function GetDebugSheet() As Worksheet
    Dim wsName As String
    wsName = "Debug_" & Format(Date, "yyyymmdd")

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(wsName)

    ' If sheet doesn't exist or error occurs, create a new one
    If ws Is Nothing Or Err.Number <> 0 Then
        ' First, try to delete if it exists but had an error
        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(wsName).Delete
        Application.DisplayAlerts = True
        Err.Clear

        ' Create new sheet and set up headers
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = wsName
        ws.Cells(1, 1).Value = "Entry"
        ws.Cells(1, 2).Value = "Time"
        ws.Cells(1, 3).Value = "Module"
        ws.Cells(1, 4).Value = "Procedure"
        ws.Cells(1, 5).Value = "Message"
        ws.Cells(1, 6).Value = "Value"

        ' Format as table for better filtering
        On Error Resume Next
        ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F1"), , xlYes).Name = "DebugTable"
        Err.Clear

        ' Make visible in case of previous settings
        ws.Visible = xlSheetVisible
    End If

    On Error GoTo 0
    Set GetDebugSheet = ws
End Function

' Simple log entry that writes directly to cells and Immediate window
Public Sub DebugLog(moduleName As String, procedureName As String, message As String, Optional value As Variant = "")
    On Error Resume Next

    ' Get debug sheet
    Dim ws As Worksheet
    Set ws = GetDebugSheet()

    ' Find next row - simple approach
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Write entry
    ws.Cells(nextRow, 1).Value = nextRow - 1  ' Entry number
    ws.Cells(nextRow, 2).Value = Now          ' Timestamp
    ws.Cells(nextRow, 3).Value = moduleName   ' Module
    ws.Cells(nextRow, 4).Value = procedureName ' Procedure
    ws.Cells(nextRow, 5).Value = message      ' Message

    ' Handle variant value
    If VarType(value) = vbObject Then
        If value Is Nothing Then
            ws.Cells(nextRow, 6).Value = "[Nothing]"
        Else
            ws.Cells(nextRow, 6).Value = "[Object: " & TypeName(value) & "]"
        End If
    ElseIf IsArray(value) Then
        ws.Cells(nextRow, 6).Value = "[Array]"
    Else
        ws.Cells(nextRow, 6).Value = value
    End If

    ' Also write to Immediate Window
    Debug.Print Format(Now, "hh:mm:ss") & " | " & moduleName & "." & procedureName & " | " & message & " | " & IIf(IsEmpty(value), "", CStr(value))

    On Error GoTo 0  ' Restore error handling
End Sub

' Examine and log information about a worksheet
Public Sub DebugSheet(ws As Worksheet, moduleName As String, procedureName As String)
    On Error Resume Next

    ' Log basic sheet properties
    DebugLog moduleName, procedureName, "Sheet Check: " & ws.Name, "Visible=" & ws.Visible
    DebugLog moduleName, procedureName, "Sheet UsedRange", ws.UsedRange.Address
    DebugLog moduleName, procedureName, "Sheet Last Cell", ws.Cells.SpecialCells(xlCellTypeLastCell).Address

    ' Check first few rows of data
    Dim i As Long, j As Long
    Dim hasData As Boolean: hasData = False

    For i = 1 To 5  ' Check first 5 rows
        For j = 1 To 5  ' Check first 5 columns
            If Len(ws.Cells(i, j).Value) > 0 Then
                DebugLog moduleName, procedureName, "Data at " & ws.Cells(i, j).Address, ws.Cells(i, j).Value
                hasData = True
            End If
        Next j
    Next i

    If Not hasData Then
        DebugLog moduleName, procedureName, "WARNING: No data found in first 5x5 cells", ""
    End If

    On Error GoTo 0
End Sub

' Get list of all sheets and their visibility
Public Sub DebugListSheets(moduleName As String, procedureName As String)
    On Error Resume Next

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        DebugLog moduleName, procedureName, "Sheet: " & ws.Name, _
                "Visible=" & Choose(ws.Visible + 1, "Visible", "Hidden", "VeryHidden")
    Next ws

    On Error GoTo 0
End Sub

' Run a quick diagnostic test of the logger itself
Public Sub SelfTest()
    DebugLog "StandaloneDebug", "SelfTest", "Starting self-test", Now
    DebugLog "StandaloneDebug", "SelfTest", "Testing string value", "Hello World"
    DebugLog "StandaloneDebug", "SelfTest", "Testing numeric value", 12345
    DebugLog "StandaloneDebug", "SelfTest", "Testing date value", Date
    DebugLog "StandaloneDebug", "SelfTest", "Testing Nothing object", Nothing

    ' List all sheets
    DebugListSheets "StandaloneDebug", "SelfTest"

    ' Test on RunLog sheet if exists
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RunLog")
    If Not ws Is Nothing Then
        DebugSheet ws, "StandaloneDebug", "SelfTest"
    End If

    DebugLog "StandaloneDebug", "SelfTest", "Self-test complete", Now
End Sub

' Simple direct single-cell write test - absolute minimum
Public Sub DirectCellWrite()
    On Error Resume Next

    Dim wsName As String: wsName = "DirectWrite"
    Dim ws As Worksheet

    ' Create sheet if needed
    If Not SheetExists(wsName) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = wsName
    End If
    Set ws = ThisWorkbook.Sheets(wsName)

    ' Write to A1
    ws.Range("A1").Value = "Direct Write Test: " & Now

    ' Show result
    If Err.Number = 0 Then
        MsgBox "Successfully wrote to cell A1 on sheet '" & wsName & "'", vbInformation
    Else
        MsgBox "ERROR: Failed to write to cell. Error #" & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not (ThisWorkbook.Sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function
