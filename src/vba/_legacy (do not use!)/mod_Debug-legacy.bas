Option Explicit

'========================================================
' Module: mod_Debug
' Contains helper subroutines for debugging purposes.
'========================================================

Sub DumpHeaders()
    ' Purpose: Prints the column order and names of the first table on the active sheet to the Immediate Window.
    Dim lo As ListObject
    Dim c As Range
    Dim ws As Worksheet

    ' Ensure the target sheet is active or specify it directly
    On Error Resume Next ' Handle error if sheet doesn't exist or no table
    Set ws = ThisWorkbook.Sheets("CurrentMonthData") ' Specify the sheet name directly
    If ws Is Nothing Then
         Debug.Print "ERROR (DumpHeaders): Sheet 'CurrentMonthData' not found."
         Exit Sub
    End If
    If ws.ListObjects.Count = 0 Then
         Debug.Print "ERROR (DumpHeaders): No ListObject (table) found on sheet 'CurrentMonthData'."
         Exit Sub
    End If
    Set lo = ws.ListObjects(1) ' Assumes the first table is the target
    On Error GoTo 0 ' Restore error handling

    Debug.Print "=== Header order from DumpHeaders (Sheet: " & ws.Name & ", Table: " & lo.Name & ") ==="
    If lo.HeaderRowRange Is Nothing Then
         Debug.Print "ERROR (DumpHeaders): Table '" & lo.Name & "' does not have a header row range."
         Exit Sub
    End If

    For Each c In lo.HeaderRowRange.Cells
        Debug.Print "Col " & c.Column & ": ", c.Value
    Next c
    Debug.Print "========================================="

End Sub


'========================================================
'  Helper - Append a message to sheet "DebugTrace"
'  *** Declared PUBLIC because it's called from other modules ***
'========================================================
Public Sub DebugTrace(tag As String, msg As String)
    ' Purpose: Writes a timestamped log entry to the DebugTrace sheet.
    '          Creates the sheet and header if it doesn't exist.
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim callerName As String

    ' Attempt to get the caller's name (might not always work reliably)
    On Error Resume Next
    callerName = Application.Caller(1) ' Get immediate caller
    If IsError(callerName) Or Len(callerName) = 0 Then
        callerName = "Unknown"
    End If
    On Error GoTo 0 ' Restore default error handling

    ' Get or create the DebugTrace sheet
    On Error Resume Next ' Handle error if sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets("DebugTrace")
    If ws Is Nothing Then ' Sheet doesn't exist, create it
        Application.ScreenUpdating = False ' Prevent screen flicker during creation
        Set ws = ThisWorkbook.Worksheets.Add(After:= _
                 ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "DebugTrace"
        ' Add headers
        ws.Range("A1:D1").Value = Array("Timestamp", "Tag", "Message", "Caller")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:D").AutoFit ' Autofit columns
        Application.ScreenUpdating = True
        Debug.Print Time & " - Created DebugTrace sheet."
    End If
    On Error GoTo 0 ' Restore default error handling

    ' Write the log entry
    If Not ws Is Nothing Then
        On Error Resume Next ' Handle potential errors writing to sheet
        With ws
            nextRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
            ' Check if sheet is empty (Header only)
            If nextRow = 1 And .Range("A1").Value = "" Then nextRow = 1 ' Handle first write after creation if A1 is empty
            If nextRow = 2 And .Range("A1").Value <> "" And .Range("A2").Value = "" Then nextRow = 2 ' Handle first data row write

            .Cells(nextRow, 1).Value = Now
            .Cells(nextRow, 1).NumberFormat = "m/d/yyyy h:mm:ss AM/PM" ' Format timestamp
            .Cells(nextRow, 2).Value = tag
            .Cells(nextRow, 3).Value = "'" & msg ' Prepend apostrophe to treat as text
            .Cells(nextRow, 4).Value = callerName
            ' Optional: Autofit column C occasionally if messages are long
            ' If nextRow Mod 20 = 0 Then .Columns("C").AutoFit
        End With
        If Err.Number <> 0 Then
             Debug.Print Time & " - ERROR writing to DebugTrace sheet: " & Err.Description
             Err.Clear
        End If
        On Error GoTo 0
    Else
         Debug.Print Time & " - ERROR: Could not get or create DebugTrace sheet."
    End If

    Set ws = Nothing ' Clean up object

End Sub


