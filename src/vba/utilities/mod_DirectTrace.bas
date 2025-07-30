' ===== mod_DirectTrace =========================================
Option Explicit

Public Sub Trace_ColumnMoves()
    Dim lo As ListObject
    Set lo = Worksheets("CurrentMonthData").ListObjects(1)

    Debug.Print String(80, "=")
    Debug.Print "TRACE – starting ReorganizeColumns on table:", lo.Name, "Time:", Now
    Debug.Print String(80, "=")

    ReorganizeColumns lo, DebugMode:=True

    Debug.Print String(80, "=")
    Debug.Print "TRACE – done.  Time:", Now
    Debug.Print String(80, "=")
End Sub
