' =========  mod_Format.bas  =========
' Purpose: Handles all presentation work for the 510(k) data table,
'          including column addition, number formatting, styling, sorting, etc.
' Key APIs exposed: AddScoreColumnsIfNeeded, ApplyAll
' Maintainer: [Your Name/Team]
' Dependencies: mod_Logger, mod_DebugTraceHelpers, mod_Schema, mod_Config
' =====================================
Option Explicit

Public Function AddScoreColumnsIfNeeded(tbl As ListObject) As Boolean
    ' Purpose: Adds the necessary scoring output columns to the table if they don't exist.
    Const PROC_NAME As String = "mod_Format.AddScoreColumnsIfNeeded"
    AddScoreColumnsIfNeeded = False ' Default to failure
    Dim colNames As Variant, colName As Variant, lc As ListColumn
    ' --- Ensure these match the columns expected by mod_Score and mod_510k_Processor ---
    colNames = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")

    On Error GoTo AddColError
    For Each colName In colNames
        On Error Resume Next ' Check if column exists
        Set lc = tbl.ListColumns(colName)
        On Error GoTo AddColError ' Restore handler

        If lc Is Nothing Then
            Set lc = tbl.ListColumns.Add
            lc.Name = colName
            LogEvt PROC_NAME, lgINFO, "Added missing column: " & colName, "Table=" & tbl.Name
            TraceEvt lvlINFO, PROC_NAME, "Added missing column", "Table=" & tbl.Name & ", Column=" & colName
        Else
            LogEvt PROC_NAME, lgDETAIL, "Column already exists: " & colName, "Table=" & tbl.Name
            TraceEvt lvlDET, PROC_NAME, "Column already exists", "Table=" & tbl.Name & ", Column=" & colName
        End If
        Set lc = Nothing ' Reset for next iteration
    Next colName

    AddScoreColumnsIfNeeded = True ' Success
    Exit Function

AddColError:
    LogEvt PROC_NAME, lgERROR, "Error adding/checking column '" & colName & "' to table '" & tbl.Name & "': " & Err.Description
    TraceEvt lvlERROR, PROC_NAME, "Error adding/checking column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "Error adding required column '" & colName & "' to table '" & tbl.Name & "': " & Err.Description, vbCritical, "Column Error"
    ' AddScoreColumnsIfNeeded remains False
End Function

Public Function ApplyAll(tbl As ListObject, wsData As Worksheet) As Boolean
    ' Purpose: Orchestrates all formatting steps for the data table.
    Const PROC_NAME As String = "mod_Format.ApplyAll"
    ApplyAll = False ' Default to failure

    If tbl Is Nothing Or wsData Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Invalid arguments (Table or Worksheet is Nothing)."
        TraceEvt lvlERROR, PROC_NAME, "Invalid arguments", "TableIsNothing=" & (tbl Is Nothing) & ", WsIsNothing=" & (wsData Is Nothing)
        Exit Function
    End If

    On Error GoTo ApplyAllError

    LogEvt PROC_NAME, lgINFO, "Starting formatting sequence for table: " & tbl.Name
    TraceEvt lvlINFO, PROC_NAME, "Start formatting sequence", "Table=" & tbl.Name

    ' --- Call individual formatting routines ---
    Call ApplyNumberFormats(tbl)
    Call FormatTableLook(tbl)
    Call FormatCategoryColors(tbl)
    Call CreateShortNamesAndComments(tbl) ' Must run after data write, before reorg/sort
    Call ReorganizeColumns(tbl)
    Call SortDataTable(tbl)
    Call FreezeHeaderAndFirstColumns(wsData)

    ApplyAll = True ' Success
    LogEvt PROC_NAME, lgINFO, "Formatting sequence completed successfully for table: " & tbl.Name
    TraceEvt lvlINFO, PROC_NAME, "Formatting sequence complete", "Table=" & tbl.Name
    Exit Function

ApplyAllError:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    LogEvt PROC_NAME, lgERROR, "Error during formatting sequence for table '" & tbl.Name & "'. Error #" & errNum & ": " & errDesc
    TraceEvt lvlERROR, PROC_NAME, "Error during formatting sequence", "Table='" & tbl.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "An error occurred during table formatting: " & vbCrLf & errDesc, vbExclamation, "Formatting Error"
    ' ApplyAll remains False
End Function

' ==========================================================================
' ===                  PRIVATE FORMATTING HELPERS                      ===
' ==========================================================================

Private Sub ApplyNumberFormats(tbl As ListObject)
    ' Purpose: Applies specific number formats to relevant columns.
    Const PROC_NAME As String = "mod_Format.ApplyNumberFormats"
    On Error GoTo FormatError
    LogEvt PROC_NAME, lgDETAIL, "Applying number formats...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Start applying number formats", "Table=" & tbl.Name

    ' Example: Format score columns as numbers with 1 decimal place
    Dim scoreCols As Variant: scoreCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score")
    Dim colName As Variant
    For Each colName In scoreCols
        On Error Resume Next ' Ignore if column doesn't exist (should have been added)
        tbl.ListColumns(colName).DataBodyRange.NumberFormat = "0.0"
        If Err.Number <> 0 Then
            LogEvt PROC_NAME, lgWARN, "Could not format column: " & colName, "Table=" & tbl.Name & ", Err=" & Err.Description
            TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Description
            Err.Clear
        End If
        On Error GoTo FormatError ' Restore handler
    Next colName

    ' Format Score_Percent as Percentage
    On Error Resume Next
    tbl.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
     If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Could not format column: Score_Percent", "Table=" & tbl.Name & ", Err=" & Err.Description
        TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=Score_Percent, Err=" & Err.Description
        Err.Clear
    End If
    On Error GoTo FormatError ' Restore handler

    ' Format Date columns
    Dim dateCols As Variant: dateCols = Array("DecisionDate", "DateReceived")
     For Each colName In dateCols
        On Error Resume Next
        tbl.ListColumns(colName).DataBodyRange.NumberFormat = "m/d/yyyy"
         If Err.Number <> 0 Then
            LogEvt PROC_NAME, lgWARN, "Could not format column: " & colName, "Table=" & tbl.Name & ", Err=" & Err.Description
            TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Description
            Err.Clear
        End If
        On Error GoTo FormatError ' Restore handler
    Next colName

    ' Format ProcTimeDays as Integer
    On Error Resume Next
    tbl.ListColumns("ProcTimeDays").DataBodyRange.NumberFormat = "0"
     If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Could not format column: ProcTimeDays", "Table=" & tbl.Name & ", Err=" & Err.Description
        TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=ProcTimeDays, Err=" & Err.Description
        Err.Clear
    End If
    On Error GoTo FormatError ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Number formats applied.", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Number formats applied", "Table=" & tbl.Name
    Exit Sub
FormatError:
    LogEvt PROC_NAME, lgERROR, "Error applying number formats: " & Err.Description, "Table=" & tbl.Name
    TraceEvt lvlERROR, PROC_NAME, "Error applying number formats", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    ' Optionally re-raise or handle, for now just log and exit sub
End Sub

Private Sub FormatTableLook(tbl As ListObject)
    ' Purpose: Applies basic table styling (e.g., style, autofit).
    Const PROC_NAME As String = "mod_Format.FormatTableLook"
    On Error Resume Next ' Be lenient with formatting errors
    LogEvt PROC_NAME, lgDETAIL, "Applying table style and autofit...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Start applying table style/autofit", "Table=" & tbl.Name

    ' Apply a standard table style (adjust name as needed)
    tbl.TableStyle = "TableStyleMedium9" ' Example style

    ' Autofit columns
    tbl.Range.Columns.AutoFit

    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Error applying table style/autofit: " & Err.Description, "Table=" & tbl.Name
        TraceEvt lvlWARN, PROC_NAME, "Error applying table style/autofit", "Table=" & tbl.Name & ", Err=" & Err.Description
        Err.Clear
    Else
        LogEvt PROC_NAME, lgDETAIL, "Table style and autofit applied.", "Table=" & tbl.Name
        TraceEvt lvlDET, PROC_NAME, "Table style/autofit applied", "Table=" & tbl.Name
    End If
    On Error GoTo 0 ' Restore default
End Sub

Private Sub FormatCategoryColors(tbl As ListObject)
    ' Purpose: Applies conditional formatting based on the 'Category' column.
    Const PROC_NAME As String = "mod_Format.FormatCategoryColors"
    Dim catCol As ListColumn, catRange As Range, cfRule As FormatCondition
    On Error GoTo FormatError
    LogEvt PROC_NAME, lgDETAIL, "Applying category conditional formatting...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Start applying category colors", "Table=" & tbl.Name

    On Error Resume Next ' Check if column exists
    Set catCol = tbl.ListColumns("Category")
    On Error GoTo FormatError ' Restore handler
    If catCol Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Category column not found. Skipping color formatting.", "Table=" & tbl.Name
        TraceEvt lvlWARN, PROC_NAME, "Category column not found", "Table=" & tbl.Name
        Exit Sub
    End If

    Set catRange = catCol.DataBodyRange
    catRange.FormatConditions.Delete ' Clear existing rules on this range

    ' --- Define Category Colors (Consider moving to mod_Config) ---
    Dim catColors As Object: Set catColors = CreateObject("Scripting.Dictionary")
    catColors("A") = RGB(198, 239, 206) ' Green Fill
    catColors("B") = RGB(255, 235, 156) ' Yellow Fill
    catColors("C") = RGB(255, 199, 206) ' Red Fill
    catColors("D") = RGB(217, 217, 217) ' Gray Fill

    Dim catKey As Variant
    For Each catKey In catColors.Keys
        Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & catKey & """")
        With cfRule.Interior
            .PatternColorIndex = xlAutomatic
            .Color = catColors(catKey)
            .TintAndShade = 0
        End With
        ' Set font color based on background brightness (Requires GetBrightness from mod_Utils)
        If mod_Utils.GetBrightness(catColors(catKey)) < 0.5 Then ' Dark background
            cfRule.Font.Color = vbWhite
        Else ' Light background
            cfRule.Font.Color = vbBlack
        End If
        cfRule.StopIfTrue = False ' Apply multiple rules if needed (though unlikely here)
    Next catKey

    LogEvt PROC_NAME, lgDETAIL, "Category conditional formatting applied.", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Category colors applied", "Table=" & tbl.Name
    Exit Sub
FormatError:
    LogEvt PROC_NAME, lgERROR, "Error applying category conditional formatting: " & Err.Description, "Table=" & tbl.Name
    TraceEvt lvlERROR, PROC_NAME, "Error applying category colors", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
End Sub

Private Sub CreateShortNamesAndComments(tbl As ListObject)
    ' Purpose: Adds comments with full names to columns with shortened headers.
    '          (Placeholder - implement specific logic if needed)
    Const PROC_NAME As String = "mod_Format.CreateShortNamesAndComments"
    On Error Resume Next ' Be lenient
    LogEvt PROC_NAME, lgDETAIL, "Applying short names/comments (Placeholder)...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Applying short names/comments (Placeholder)", "Table=" & tbl.Name

    ' --- Example Placeholder Logic ---
    ' Dim lc As ListColumn
    ' For Each lc In tbl.ListColumns
    '     Select Case lc.Name
    '         Case "AC_Wt": lc.Range.Cells(1).AddComment "Advisory Committee Weight"
    '         Case "PC_Wt": lc.Range.Cells(1).AddComment "Product Code Weight"
    '         ' Add other cases as needed
    '     End Select
    ' Next lc
    ' --- End Placeholder ---

    LogEvt PROC_NAME, lgDETAIL, "Short names/comments applied (Placeholder).", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Short names/comments applied (Placeholder)", "Table=" & tbl.Name
    On Error GoTo 0
End Sub

Private Sub ReorganizeColumns(tbl As ListObject)
    ' Purpose: Moves columns to a predefined order.
    Const PROC_NAME As String = "mod_Format.ReorganizeColumns"
    Dim desiredOrder As Variant, currentPos As Long, targetPos As Long, colName As Variant, lc As ListColumn
    ' --- Define Desired Order (Consider moving to mod_Config) ---
    desiredOrder = Array("K_Number", "Applicant", "DeviceName", "Category", "Final_Score", "Score_Percent", "CompanyRecap", "DecisionDate", "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "Country", "Statement", "FDA_Link", "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc")

    On Error GoTo ReorgError
    LogEvt PROC_NAME, lgDETAIL, "Reorganizing columns...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Start reorganizing columns", "Table=" & tbl.Name

    Application.ScreenUpdating = False ' Speed up column moves

    targetPos = 1
    For Each colName In desiredOrder
        On Error Resume Next ' Check if column exists
        Set lc = tbl.ListColumns(colName)
        On Error GoTo ReorgError ' Restore handler

        If Not lc Is Nothing Then
            currentPos = lc.Index
            If currentPos <> targetPos Then
                lc.Range.EntireColumn.Cut
                tbl.HeaderRowRange.Parent.Columns(targetPos).Insert Shift:=xlToRight
                Application.CutCopyMode = False ' Clear clipboard
                LogEvt PROC_NAME, lgDETAIL, "Moved column '" & colName & "' from " & currentPos & " to " & targetPos, "Table=" & tbl.Name
                TraceEvt lvlDET, PROC_NAME, "Moved column", "Table=" & tbl.Name & ", Col=" & colName & ", From=" & currentPos & ", To=" & targetPos
            End If
            targetPos = targetPos + 1
        Else
            LogEvt PROC_NAME, lgWARN, "Column '" & colName & "' not found for reorganization.", "Table=" & tbl.Name
            TraceEvt lvlWARN, PROC_NAME, "Column not found for reorg", "Table=" & tbl.Name & ", Col=" & colName
        End If
        Set lc = Nothing
    Next colName

    Application.ScreenUpdating = True
    LogEvt PROC_NAME, lgDETAIL, "Column reorganization complete.", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Column reorganization complete", "Table=" & tbl.Name
    Exit Sub

ReorgError:
    Application.ScreenUpdating = True ' Ensure screen updating is back on
    LogEvt PROC_NAME, lgERROR, "Error reorganizing columns: " & Err.Description, "Table=" & tbl.Name
    TraceEvt lvlERROR, PROC_NAME, "Error reorganizing columns", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "An error occurred while reorganizing columns: " & Err.Description, vbExclamation, "Column Reorganization Error"
End Sub

Private Sub SortDataTable(tbl As ListObject)
    ' Purpose: Sorts the table by the primary sort key(s).
    Const PROC_NAME As String = "mod_Format.SortDataTable"
    Dim sortCol As Range
    On Error GoTo SortError
    LogEvt PROC_NAME, lgDETAIL, "Sorting data table...", "Table=" & tbl.Name
    TraceEvt lvlDET, PROC_NAME, "Start sorting table", "Table=" & tbl.Name

    ' --- Define Sort Key (Consider moving to mod_Config) ---
    Const SORT_COLUMN_NAME As String = "Final_Score"
    Const SORT_ORDER As XlSortOrder = xlDescending

    On Error Resume Next ' Check if sort column exists
    Set sortCol = tbl.ListColumns(SORT_COLUMN_NAME).Range
    On Error GoTo SortError ' Restore handler

    If sortCol Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Sort column '" & SORT_COLUMN_NAME & "' not found. Skipping sort.", "Table=" & tbl.Name
        TraceEvt lvlWARN, PROC_NAME, "Sort column not found", "Table=" & tbl.Name & ", Col=" & SORT_COLUMN_NAME
        Exit Sub
    End If

    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortCol, SortOn:=xlSortOnValues, Order:=SORT_ORDER, DataOption:=xlSortNormal
        ' Add secondary sort keys if needed:
        ' .SortFields.Add Key:=tbl.ListColumns("SecondaryColumn").Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    LogEvt PROC_NAME, lgDETAIL, "Data table sorted.", "Table=" & tbl.Name & ", Key=" & SORT_COLUMN_NAME
    TraceEvt lvlDET, PROC_NAME, "Table sorted", "Table=" & tbl.Name & ", Key=" & SORT_COLUMN_NAME
    Exit Sub

SortError:
    LogEvt PROC_NAME, lgERROR, "Error sorting table: " & Err.Description, "Table=" & tbl.Name
    TraceEvt lvlERROR, PROC_NAME, "Error sorting table", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "An error occurred while sorting the table: " & Err.Description, vbExclamation, "Sort Error"
End Sub

Private Sub FreezeHeaderAndFirstColumns(ws As Worksheet)
    ' Purpose: Freezes the header row and the first few columns for better navigation.
    Const PROC_NAME As String = "mod_Format.FreezeHeaderAndFirstColumns"
    Const COLUMNS_TO_FREEZE As Long = 3 ' e.g., K_Number, Applicant, DeviceName
    On Error Resume Next ' Be lenient with UI operations
    LogEvt PROC_NAME, lgDETAIL, "Freezing panes...", "Sheet=" & ws.Name
    TraceEvt lvlDET, PROC_NAME, "Start freezing panes", "Sheet=" & ws.Name

    ws.Activate ' Must activate sheet to set freeze panes
    ActiveWindow.FreezePanes = False ' Unfreeze first
    With ws.Cells(2, COLUMNS_TO_FREEZE + 1) ' Cell below header, right of columns to freeze
        .Activate ' Select the cell to set the freeze boundary
        ActiveWindow.FreezePanes = True
    End With
    ws.Range("A1").Activate ' Select A1 after freezing

    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Error freezing panes: " & Err.Description, "Sheet=" & ws.Name
        TraceEvt lvlWARN, PROC_NAME, "Error freezing panes", "Sheet=" & ws.Name & ", Err=" & Err.Description
        Err.Clear
    Else
        LogEvt PROC_NAME, lgDETAIL, "Panes frozen.", "Sheet=" & ws.Name
        TraceEvt lvlDET, PROC_NAME, "Panes frozen", "Sheet=" & ws.Name
    End If
    On Error GoTo 0 ' Restore default
End Sub
