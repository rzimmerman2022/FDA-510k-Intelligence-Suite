' ===== mod_ColumnDebugger =========================================
Option Explicit

'-----------------------------------------------------------------------------
' Comprehensive debugging tool to diagnose column move issues
' - Prints exact column names and positions
' - Verifies if column names match what's expected
' - Monitors the table edge position
' - Creates detailed trace of the move process
'-----------------------------------------------------------------------------
Public Sub Debug_ColumnMoves()
    On Error Resume Next
    Dim lo As ListObject
    Set lo = Worksheets("CurrentMonthData").ListObjects(1)
    
    If lo Is Nothing Then
        Debug.Print "ERROR: Could not find ListObject on CurrentMonthData sheet"
        Exit Sub
    End If
    
    Debug.Print String(80, "=")
    Debug.Print "COLUMN DEBUGGER - Table: " & lo.Name & " on " & lo.Parent.Name & " at " & Now
    Debug.Print String(80, "=")
    
    ' Report the absolute position of the table's left edge
    Dim firstTableCol As Long
    firstTableCol = lo.Range.Column
    Debug.Print "Table's left edge (firstTableCol) is at column " & firstTableCol & " (" & _
                ColumnLetterFromNumber(firstTableCol) & ")"
    
    ' Print exact column order with detailed position info
    DumpColumnDetails lo
    
    ' List desired target positions from ReorganizeColumns
    DumpTargetOrder
    
    ' Verify exact case/spelling of some key columns we're trying to move
    VerifyColumnNames lo, Array("DeviceName", "Contact", "Applicant", "Final_Score", "FDA_Link", "Category")
    
    ' Run ReorganizeColumns with debugging on
    Debug.Print String(80, "-")
    Debug.Print "Running ReorganizeColumns with DebugMode=True..."
    ReorganizeColumns lo, DebugMode:=True
    Debug.Print String(80, "-")
    
    ' Print final column position after moves
    Debug.Print "AFTER ReorganizeColumns - final column positions:"
    DumpColumnDetails lo
    
    Debug.Print String(80, "=")
    Debug.Print "COLUMN DEBUGGER FINISHED at " & Now
    Debug.Print String(80, "=")
End Sub

'-----------------------------------------------------------------------------
' Helper functions
'-----------------------------------------------------------------------------

' Print detailed information about each column in the ListObject
Private Sub DumpColumnDetails(lo As ListObject)
    Debug.Print String(60, "-")
    Debug.Print "COLUMN DETAILS:"
    Debug.Print "Index" & vbTab & "Abs.Col" & vbTab & vbTab & "Name"
    Debug.Print String(60, "-")
    
    Dim i As Long, lc As ListColumn
    Dim absCol As Long
    Dim firstTableCol As Long: firstTableCol = lo.Range.Column
    
    For i = 1 To lo.ListColumns.Count
        Set lc = lo.ListColumns(i)
        absCol = firstTableCol + i - 1  ' Convert to absolute sheet column
        Debug.Print i & vbTab & absCol & " (" & ColumnLetterFromNumber(absCol) & ")" & vbTab & """" & lc.Name & """"
    Next i
    
    Debug.Print String(60, "-")
End Sub

' Print the target order from ReorganizeColumns
Private Sub DumpTargetOrder()
    Debug.Print String(60, "-")
    Debug.Print "TARGET ORDER from ReorganizeColumns array:"
    Debug.Print "Pos" & vbTab & "Column Name"
    Debug.Print String(60, "-")
    
    Dim targetOrder As Variant
    targetOrder = Array( _
        "K_Number", "DecisionDate", "Applicant", "DeviceName", _
        "Contact", "CompanyRecap", "Score_Percent", "Category", _
        "FDA_Link", "Final_Score", "DateReceived", "ProcTimeDays", _
        "AC", "PC", "SubmType", "Country", "Statement", _
        "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", _
        "GL_Wt", "NF_Calc", "Synergy_Calc", "City", "State")
    
    Dim i As Long
    For i = LBound(targetOrder) To UBound(targetOrder)
        Debug.Print (i + 1) & vbTab & """" & targetOrder(i) & """"
    Next i
    
    Debug.Print String(60, "-")
End Sub

' Verify if the specified column names exist in the table (exact match)
Private Sub VerifyColumnNames(lo As ListObject, colNames As Variant)
    Debug.Print String(60, "-")
    Debug.Print "COLUMN NAME VERIFICATION:"
    Debug.Print String(60, "-")
    
    Dim colName As Variant, lc As ListColumn
    
    For Each colName In colNames
        On Error Resume Next
        Set lc = lo.ListColumns(CStr(colName))
        Dim errNum As Long: errNum = Err.Number
        On Error GoTo 0
        
        If errNum = 0 And Not lc Is Nothing Then
            Debug.Print "✓ Column '" & colName & "' EXISTS at index " & lc.Index
        Else
            Debug.Print "✗ Column '" & colName & "' NOT FOUND or ERROR"
            ' Try to find similar columns in case of case sensitivity issues
            FindSimilarColumns lo, CStr(colName)
        End If
    Next colName
    
    Debug.Print String(60, "-")
End Sub

' Try to find columns with similar names (case insensitive)
Private Sub FindSimilarColumns(lo As ListObject, searchName As String)
    Dim i As Long, lc As ListColumn
    Debug.Print "  Searching for similar names to '" & searchName & "':"
    
    Dim found As Boolean: found = False
    For i = 1 To lo.ListColumns.Count
        Set lc = lo.ListColumns(i)
        If LCase(lc.Name) = LCase(searchName) Then
            Debug.Print "  → Found '" & lc.Name & "' at index " & i & " (case mismatch)"
            found = True
        ElseIf InStr(1, lc.Name, searchName, vbTextCompare) > 0 Then
            Debug.Print "  → Similar: '" & lc.Name & "' at index " & i
            found = True
        End If
    Next i
    
    If Not found Then Debug.Print "  → No similar names found"
End Sub

' Convert column number to Excel column letter (A, B, ..., AA, AB, etc.)
Private Function ColumnLetterFromNumber(colNum As Long) As String
    Dim result As String
    Dim modVal As Integer
    
    Do While colNum > 0
        modVal = (colNum - 1) Mod 26
        result = Chr(65 + modVal) & result
        colNum = (colNum - modVal - 1) \ 26
    Loop
    
    ColumnLetterFromNumber = result
End Function
