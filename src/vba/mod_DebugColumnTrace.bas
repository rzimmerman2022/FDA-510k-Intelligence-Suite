' ===== mod_DebugColumnTrace =========================================
Option Explicit

Private Sub DumpOrder(lo As ListObject, label As String)
    Dim i As Long, out$
    For i = 1 To lo.ListColumns.Count
        out = out & lo.ListColumns(i).Name & " | "
    Next i
    Debug.Print label & " (" & lo.Name & "):"
    Debug.Print Left$(out, Len(out) - 3)
End Sub

Public Sub Trace_ColumnMoves()
    Dim ws As Worksheet, lo As ListObject
    Set ws = ActiveSheet
    If ws.ListObjects.Count = 0 Then
        MsgBox "Active sheet has no table.", vbExclamation: Exit Sub
    End If
    Set lo = ws.ListObjects(1)
    
    Debug.Print String(70, "=")
    Debug.Print "TRACE START – sheet: " & ws.Name & ", table: " & lo.Name
    Debug.Print "Total comments BEFORE run: "; ActiveSheet.Comments.Count
    DumpOrder lo, "Initial order"
    
    '--- run ReorganizeColumns with DebugMode:=True so it prints every move
    Debug.Print String(50, "-")
    ReorganizeColumns lo, DebugMode:=True
    Debug.Print String(50, "-")
    
    DumpOrder lo, "After ReorganizeColumns"
    Debug.Print "Total comments AFTER Reorg: "; ActiveSheet.Comments.Count
    
    '--- now create notes (short triangles)
    CreateShortNamesAndComments lo, maxLen:=40, DebugMode:=True
    Debug.Print "Total comments AFTER CreateShortNamesAndComments: "; _
                ActiveSheet.Comments.Count
    Debug.Print String(70, "=")
End Sub
