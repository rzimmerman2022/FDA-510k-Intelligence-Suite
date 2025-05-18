' ==========================================================================
' Module      : mod_Format
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module is responsible for applying all visual formatting
'               and presentation adjustments to the main data table after
'               scoring is complete. This includes adding necessary output
'               columns, applying number formats, setting table styles,
'               applying conditional formatting (e.g., category colors),
'               reorganizing columns into a standard layout, sorting the
'               data, and freezing panes for usability.
'
' Key Functions:
'               - AddScoreColumnsIfNeeded: Ensures all required output columns
'                 exist in the table before processing.
'               - ApplyAll: Orchestrates the entire formatting sequence by
'                 calling various private helper subroutines.
'
' Private Helpers:
'               - ApplyNumberFormats: Sets formats like percentage, date, number.
'               - FormatTableLook: Applies Excel table styles and autofits columns.
'               - FormatCategoryColors: Applies conditional background/font colors
'                 based on the 'Category' column value.
'               - CreateShortNamesAndComments: (Placeholder) Intended for adding
'                 comments to headers with shortened names.
'               - ReorganizeColumns: Arranges columns into a predefined order.
'               - SortDataTable: Sorts the table based on defined key(s).
'               - FreezeHeaderAndFirstColumns: Freezes top row and initial columns.
'
' Dependencies: - mod_Logger: For logging events and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - mod_Schema: Used indirectly via column name constants/lookups.
'               - mod_Config: For constants like sort keys or desired column order
'                 (though some are currently hardcoded).
'               - mod_Utils: For utility functions like GetBrightness.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Corrected "Syntax error" in SortDataTable by changing
'                               Const SORT_ORDER As XlSortOrder to Dim SORT_ORDER
'                               and assigning value separately. Typed Consts are
'                               not allowed within procedures.
' 2025-04-30  Cline (AI)      - Added DeleteDuplicateColumns sub and call in ApplyAll.
' 2025-04-30  Cline (AI)      - Updated desired column order in ReorganizeColumns.
' 2025-04-30  Cline (AI)      - Commented out FreezeHeaderAndFirstColumns call in ApplyAll.
' 2025-04-30  Cline (AI)      - Added final AutoFit call in ApplyAll.
' 2025-04-30  Cline (AI)      - Corrected column order in ReorganizeColumns per user request. Ensured pane freezing remains disabled.
' 2025-04-30  Cline (AI)      - Added logging to ApplyAll steps, removed error suppression on final AutoFit.
' 2025-04-30  Cline (AI)      - Added more detailed logging within DeleteDuplicateColumns to diagnose deletion failures.
' 2025-04-30  Cline (AI)      - Modified DeleteDuplicateColumns to use pattern matching
'                               (Like "* [0-9]") instead of a hardcoded list.
' [Previous dates/authors/changes unknown]
' ==========================================================================
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
            mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Added missing column", "Table=" & tbl.Name & ", Column=" & colName
        Else
            LogEvt PROC_NAME, lgDETAIL, "Column already exists: " & colName, "Table=" & tbl.Name
            mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Column already exists", "Table=" & tbl.Name & ", Column=" & colName
        End If
        Set lc = Nothing ' Reset for next iteration
    Next colName

    AddScoreColumnsIfNeeded = True ' Success
    Exit Function

AddColError:
    LogEvt PROC_NAME, lgERROR, "Error adding/checking column '" & colName & "' to table '" & tbl.Name & "': " & Err.Description
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error adding/checking column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "Error adding required column '" & colName & "' to table '" & tbl.Name & "': " & Err.Description, vbCritical, "Column Error"
    ' AddScoreColumnsIfNeeded remains False
End Function

Public Function ApplyAll(tbl As ListObject, wsData As Worksheet) As Boolean
    ' Purpose: Orchestrates all formatting steps for the data table.
    Const PROC_NAME As String = "mod_Format.ApplyAll"
    ApplyAll = False ' Default to failure
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Entered ApplyAll", "Table=" & IIf(tbl Is Nothing, "[Nothing]", tbl.Name) ' <<< STANDALONE DEBUG

    If tbl Is Nothing Or wsData Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Invalid arguments (Table or Worksheet is Nothing)."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Invalid arguments", "TableIsNothing=" & (tbl Is Nothing) & ", WsIsNothing=" & (wsData Is Nothing)
        StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ERROR: Invalid arguments", "TableIsNothing=" & (tbl Is Nothing) ' <<< STANDALONE DEBUG
        Exit Function
    End If

    On Error GoTo ApplyAllError

    LogEvt PROC_NAME, lgINFO, "Starting formatting sequence for table: " & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start formatting sequence", "Table=" & tbl.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Starting formatting sequence", tbl.Name ' <<< STANDALONE DEBUG

    ' --- Call individual formatting routines ---
    LogEvt PROC_NAME, lgDETAIL, "Calling DeleteDuplicateColumns..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling DeleteDuplicateColumns" ' <<< STANDALONE DEBUG
    Call DeleteDuplicateColumns(tbl) ' Delete known duplicates first

    LogEvt PROC_NAME, lgDETAIL, "Calling ApplyNumberFormats..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling ApplyNumberFormats" ' <<< STANDALONE DEBUG
    Call ApplyNumberFormats(tbl)

    LogEvt PROC_NAME, lgDETAIL, "Calling FormatTableLook..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling FormatTableLook" ' <<< STANDALONE DEBUG
    Call FormatTableLook(tbl) ' Includes initial AutoFit and specific widths

    LogEvt PROC_NAME, lgDETAIL, "Calling FormatCategoryColors..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling FormatCategoryColors" ' <<< STANDALONE DEBUG
    Call FormatCategoryColors(tbl) ' Apply colors based on Category column name

    LogEvt PROC_NAME, lgDETAIL, "Calling CreateShortNamesAndComments..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling CreateShortNamesAndComments" ' <<< STANDALONE DEBUG
    Call CreateShortNamesAndComments(tbl) ' Call with default parameters

    ' --- Add logging before ReorganizeColumns ---
    LogEvt PROC_NAME, lgDETAIL, "Checking 'Category' column existence BEFORE ReorganizeColumns...", "Exists=" & ColumnExists(tbl, "Category")
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Check BEFORE Reorg", "Category Exists=" & ColumnExists(tbl, "Category")
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check BEFORE Reorg", "Category Exists=" & ColumnExists(tbl, "Category") ' <<< STANDALONE DEBUG
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check BEFORE Reorg", "Contact Exists=" & ColumnExists(tbl, "Contact") ' <<< STANDALONE DEBUG
    ' --- End logging ---

    LogEvt PROC_NAME, lgDETAIL, "Calling ReorganizeColumns..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling ReorganizeColumns" ' <<< STANDALONE DEBUG
    Call ReorganizeColumns(tbl) ' Move columns to desired order

    ' --- Add logging after ReorganizeColumns ---
    LogEvt PROC_NAME, lgDETAIL, "Checking 'Category' column existence AFTER ReorganizeColumns...", "Exists=" & ColumnExists(tbl, "Category")
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Check AFTER Reorg", "Category Exists=" & ColumnExists(tbl, "Category")
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check AFTER Reorg", "Category Exists=" & ColumnExists(tbl, "Category") ' <<< STANDALONE DEBUG
    ' Add logging for column indices using the new helper function
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check AFTER Reorg", "Contact Index=" & GetColumnIndex(tbl, "Contact") ' <<< STANDALONE DEBUG
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check AFTER Reorg", "DeviceName Index=" & GetColumnIndex(tbl, "DeviceName") ' <<< STANDALONE DEBUG
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Check AFTER Reorg", "Category Index=" & GetColumnIndex(tbl, "Category") ' <<< STANDALONE DEBUG
    ' --- End logging ---

    LogEvt PROC_NAME, lgDETAIL, "Calling SortDataTable..."
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Calling SortDataTable" ' <<< STANDALONE DEBUG
    Call SortDataTable(tbl)
    ' Call FreezeHeaderAndFirstColumns(wsData) ' Commented out as requested

    ' --- Final Autofit removed as it overrides specific widths set in FormatTableLook ---
    ' LogEvt PROC_NAME, lgDETAIL, "Attempting final AutoFit..."
    ' StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting final AutoFit" ' <<< STANDALONE DEBUG
    ' tbl.Range.Columns.AutoFit
    ' LogEvt PROC_NAME, lgDETAIL, "Final AutoFit completed."
    ' StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Final AutoFit completed" ' <<< STANDALONE DEBUG

    ApplyAll = True ' Success
    LogEvt PROC_NAME, lgINFO, "Formatting sequence completed successfully for table: " & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Formatting sequence complete", "Table=" & tbl.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ApplyAll completed successfully", tbl.Name ' <<< STANDALONE DEBUG
    Exit Function

ApplyAllError:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    LogEvt PROC_NAME, lgERROR, "Error during formatting sequence for table '" & tbl.Name & "'. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during formatting sequence", "Table='" & tbl.Name & "', Err=" & errNum & " - " & errDesc
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ERROR in ApplyAll", "Err=" & errNum & " - " & errDesc ' <<< STANDALONE DEBUG
    MsgBox "An error occurred during table formatting: " & vbCrLf & errDesc, vbExclamation, "Formatting Error"
    ' ApplyAll remains False
End Function

' ==========================================================================
' ===                  PRIVATE FORMATTING HELPERS                      ===
' ==========================================================================

Private Sub DeleteDuplicateColumns(tbl As ListObject)
    ' Purpose: Deletes duplicate columns identified by a pattern (ending in space + number).
    Const PROC_NAME As String = "mod_Format.DeleteDuplicateColumns"
    Dim lc As ListColumn
    Dim i As Long ' Use Long for column index

    LogEvt PROC_NAME, lgDETAIL, "Checking for pattern-based duplicate columns to delete...", "Table=" & tbl.Name & ", Pattern='* [0-9]'"
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start checking for pattern-based duplicate columns", "Table=" & tbl.Name

    ' Loop backwards when deleting columns to avoid index issues
    For i = tbl.ListColumns.Count To 1 Step -1
        Set lc = tbl.ListColumns(i)

        ' Check if column name matches the pattern (e.g., "Column Name 2")
        If lc.Name Like "* [0-9]" Then
            LogEvt PROC_NAME, lgDETAIL, "Pattern match found. Attempting to delete duplicate column: '" & lc.Name & "'", "Index=" & i
            mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Pattern match found, attempting delete", "Column=" & lc.Name & ", Index=" & i

            Dim errNum As Long, errDesc As String
            On Error Resume Next ' Handle error during delete
            lc.Delete
            errNum = Err.Number ' Capture error info IMMEDIATELY
            errDesc = Err.Description
            On Error GoTo 0 ' Restore default error handling NOW

            If errNum <> 0 Then
                LogEvt PROC_NAME, lgWARN, "FAILED to delete duplicate column: '" & lc.Name & "'", "Table=" & tbl.Name & ", ErrNo=" & errNum & ", ErrDesc=" & errDesc
                mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "FAILED to delete duplicate column", "Table=" & tbl.Name & ", Column=" & lc.Name & ", ErrNo=" & errNum & ", ErrDesc=" & errDesc
            Else
                LogEvt PROC_NAME, lgINFO, "Successfully deleted duplicate column: '" & lc.Name & "'", "Table=" & tbl.Name
                mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Successfully deleted duplicate column", "Table=" & tbl.Name & ", Column=" & lc.Name
            End If
            ' No Exit For needed as we check each column index independently
        End If
        Set lc = Nothing
    Next i

    LogEvt PROC_NAME, lgDETAIL, "Pattern-based duplicate column check complete.", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Pattern-based duplicate column check complete", "Table=" & tbl.Name
End Sub


Private Sub ApplyNumberFormats(tbl As ListObject)
    ' Purpose: Applies specific number formats to relevant columns.
    Const PROC_NAME As String = "mod_Format.ApplyNumberFormats"
    On Error GoTo FormatError
    LogEvt PROC_NAME, lgDETAIL, "Applying number formats...", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start applying number formats", "Table=" & tbl.Name

    ' Example: Format score columns as numbers with 1 decimal place
    Dim scoreCols As Variant: scoreCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score")
    Dim colName As Variant
    For Each colName In scoreCols
        On Error Resume Next ' Ignore if column doesn't exist (should have been added)
        tbl.ListColumns(colName).DataBodyRange.NumberFormat = "0.0"
        If Err.Number <> 0 Then
            LogEvt PROC_NAME, lgWARN, "Could not format column: " & colName, "Table=" & tbl.Name & ", Err=" & Err.Description
            mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Description
            Err.Clear
        End If
        On Error GoTo FormatError ' Restore handler
    Next colName

    ' Format Score_Percent as Percentage
    On Error Resume Next
    tbl.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
     If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Could not format column: Score_Percent", "Table=" & tbl.Name & ", Err=" & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=Score_Percent, Err=" & Err.Description
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
            mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=" & colName & ", Err=" & Err.Description
            Err.Clear
        End If
        On Error GoTo FormatError ' Restore handler
    Next colName

    ' Format ProcTimeDays as Integer
    On Error Resume Next
    tbl.ListColumns("ProcTimeDays").DataBodyRange.NumberFormat = "0"
     If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Could not format column: ProcTimeDays", "Table=" & tbl.Name & ", Err=" & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Could not format column", "Table=" & tbl.Name & ", Column=ProcTimeDays, Err=" & Err.Description
        Err.Clear
    End If
    On Error GoTo FormatError ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Number formats applied.", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Number formats applied", "Table=" & tbl.Name
    Exit Sub
FormatError:
    LogEvt PROC_NAME, lgERROR, "Error applying number formats: " & Err.Description, "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error applying number formats", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    ' Optionally re-raise or handle, for now just log and exit sub
End Sub

Private Sub FormatTableLook(tbl As ListObject)
    ' Purpose: Applies basic table styling (e.g., style, autofit).
    Const PROC_NAME As String = "mod_Format.FormatTableLook"
    On Error Resume Next ' Be lenient with formatting errors
    LogEvt PROC_NAME, lgDETAIL, "Applying table style and autofit...", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start applying table style/autofit", "Table=" & tbl.Name

    ' Apply a standard table style (adjust name as needed)
    tbl.TableStyle = "TableStyleMedium9" ' Example style

    ' Autofit columns first
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting AutoFit" ' <<< STANDALONE DEBUG
    tbl.Range.Columns.AutoFit
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "AutoFit completed" ' <<< STANDALONE DEBUG

    ' Set specific widths AFTER autofit
    On Error Resume Next ' Ignore errors if columns don't exist
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting to set Applicant width" ' <<< STANDALONE DEBUG
    tbl.ListColumns("Applicant").Range.ColumnWidth = 25 ' Narrower width
    If Err.Number = 0 Then tbl.ListColumns("Applicant").Range.EntireColumn.Hidden = False ' Attempt persistence
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Set Applicant Width", IIf(Err.Number <> 0, "ERROR: " & Err.Description, 25) ' <<< STANDALONE DEBUG
    Err.Clear ' Clear potential error before next step

    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting to set DeviceName width" ' <<< STANDALONE DEBUG
    tbl.ListColumns("DeviceName").Range.ColumnWidth = 40 ' Narrower width
    If Err.Number = 0 Then tbl.ListColumns("DeviceName").Range.EntireColumn.Hidden = False ' Attempt persistence
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Set DeviceName Width", IIf(Err.Number <> 0, "ERROR: " & Err.Description, 40) ' <<< STANDALONE DEBUG
    Err.Clear ' Clear potential error before next step

    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting to set CompanyRecap width" ' <<< STANDALONE DEBUG
    tbl.ListColumns("CompanyRecap").Range.ColumnWidth = 40
    If Err.Number = 0 Then tbl.ListColumns("CompanyRecap").Range.EntireColumn.Hidden = False ' Attempt persistence
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Set CompanyRecap Width", IIf(Err.Number <> 0, "ERROR: " & Err.Description, 40) ' <<< STANDALONE DEBUG
    Err.Clear ' Clear potential error before next step

    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting to set Category width" ' <<< STANDALONE DEBUG
    tbl.ListColumns("Category").Range.ColumnWidth = 15
    If Err.Number = 0 Then tbl.ListColumns("Category").Range.EntireColumn.Hidden = False ' Attempt persistence
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Set Category Width", IIf(Err.Number <> 0, "ERROR: " & Err.Description, 15) ' <<< STANDALONE DEBUG
    Err.Clear ' Clear potential error

    ' Add others if needed
    On Error GoTo 0 ' Restore default error handling for the rest of the sub

    ' Check for errors accumulated during width setting
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Error applying table style/autofit: " & Err.Description, "Table=" & tbl.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Error applying table style/autofit", "Table=" & tbl.Name & ", Err=" & Err.Description
        Err.Clear
    Else
        LogEvt PROC_NAME, lgDETAIL, "Table style and autofit applied.", "Table=" & tbl.Name
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Table style/autofit applied", "Table=" & tbl.Name
    End If
    On Error GoTo 0 ' Restore default
End Sub

Private Sub FormatCategoryColors(tblData As ListObject) ' Changed parameter name to tblData
    ' Purpose: Applies conditional formatting based on the 'Category' column.
    Const PROC_NAME As String = "mod_Format.FormatCategoryColors"
    Dim catCol As ListColumn, catRange As Range, cfRule As FormatCondition

    On Error GoTo FormatError
    LogEvt PROC_NAME, lgDETAIL, "Entered FormatCategoryColors", "Table=" & tblData.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Entered FormatCategoryColors", "Table=" & tblData.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Entered FormatCategoryColors", tblData.Name ' <<< STANDALONE DEBUG

    On Error Resume Next ' Check if column exists
    Set catCol = tblData.ListColumns("Category")
    On Error GoTo FormatError ' Restore handler

    If catCol Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "'Category' column not found. Skipping color formatting.", "Table=" & tblData.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "'Category' column not found", "Table=" & tblData.Name
        StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Category column NOT FOUND. Skipping.", tblData.Name ' <<< STANDALONE DEBUG
        Exit Sub ' Cannot apply formatting if the target column doesn't exist
    Else
        LogEvt PROC_NAME, lgDETAIL, "Category column FOUND. Proceeding.", "Table=" & tblData.Name
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Category column FOUND", "Table=" & tblData.Name
        StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Category column FOUND. Proceeding.", tblData.Name ' <<< STANDALONE DEBUG
    End If

    Set catRange = catCol.DataBodyRange
    catRange.FormatConditions.Delete ' Clear existing rules on this range

    ' --- Define Category Colors and Apply Formatting ---
    ' High - Green
    Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""High""")
    With cfRule.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(198, 239, 206) ' Light green
        .TintAndShade = 0
    End With
    cfRule.Font.Bold = True
    cfRule.Font.Color = RGB(0, 97, 0) ' Dark green text for contrast
    cfRule.StopIfTrue = False

    ' Moderate - Yellow
    Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Moderate""")
    With cfRule.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 235, 156) ' Light yellow
        .TintAndShade = 0
    End With
    cfRule.Font.Bold = True
    cfRule.Font.Color = RGB(156, 101, 0) ' Dark gold/brown text
    cfRule.StopIfTrue = False

    ' Low - Light Orange
    Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Low""")
    With cfRule.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 221, 179) ' Light orange
        .TintAndShade = 0
    End With
    cfRule.Font.Bold = True
    cfRule.Font.Color = RGB(156, 87, 0) ' Dark orange/brown text
    cfRule.StopIfTrue = False

    ' Almost None - Gray
    Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Almost None""")
    With cfRule.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(242, 242, 242) ' Light gray
        .TintAndShade = 0
    End With
    cfRule.Font.Bold = True
    cfRule.Font.Color = RGB(89, 89, 89) ' Dark gray text
    cfRule.StopIfTrue = False

    ' Error - Red
    Set cfRule = catRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Error""")
    With cfRule.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 199, 206) ' Light red
        .TintAndShade = 0
    End With
    cfRule.Font.Bold = True
    cfRule.Font.Color = RGB(156, 0, 6) ' Dark red text
    cfRule.StopIfTrue = False

    LogEvt PROC_NAME, lgDETAIL, "Category conditional formatting applied successfully.", "Table=" & tblData.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Category formatting applied", "Table=" & tblData.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Category formatting applied", tblData.Name ' <<< STANDALONE DEBUG
    Exit Sub

FormatError:
    LogEvt PROC_NAME, lgERROR, "Error applying category conditional formatting: " & Err.Description, "Table=" & tblData.Name
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error applying category formatting", "Table=" & tblData.Name & ", Err=" & Err.Number & " - " & Err.Description
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ERROR in FormatCategoryColors", "Err=" & Err.Number & " - " & Err.Description ' <<< STANDALONE DEBUG
    Debug.Print Time & " - ERROR in " & PROC_NAME & ": " & Err.Description
End Sub

'-----------------------------------------------------------------
' Show a 40-char visible name; add a legacy NOTE (red triangle)
' ***only*** when we actually truncated the text.
' Added enhanced debugging to trace exactly what's happening.
'-----------------------------------------------------------------
Public Sub CreateShortNamesAndComments( _
        lo As ListObject, _
        Optional maxLen As Long = 40, _
        Optional tgtColName As String = "DeviceName", _
        Optional DebugMode As Boolean = True)  ' Default to debug ON

    Const PROC As String = "CreateShortNamesAndComments"
    On Error GoTo ErrHandler

    ' Debug early failures clearly
    If lo Is Nothing Then
        Debug.Print PROC & " - Error: ListObject is Nothing!"
        LogEvt PROC, lgERROR, "ListObject is Nothing!"
        Exit Sub
    End If
    
    Debug.Print PROC & " - Working with table: " & lo.Name
    
    ' Debug column existence
    Debug.Print PROC & " - Looking for column: " & tgtColName
    
    Dim tgtCol As ListColumn
    On Error Resume Next
    Set tgtCol = lo.ListColumns(tgtColName)
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    On Error GoTo ErrHandler
    
    If tgtCol Is Nothing Then
        Debug.Print PROC & " - Error: Column """ & tgtColName & """ not found! Err=" & errNum & " - " & errDesc
        LogEvt PROC, lgERROR, "Column not found: " & tgtColName & ", Err=" & errNum & " - " & errDesc
        Exit Sub
    Else
        Debug.Print PROC & " - Found column: " & tgtColName & " at index: " & tgtCol.Index
    End If

    ' Show red triangles only (keeps sheet clean)
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    Debug.Print PROC & " - Set DisplayCommentIndicator to xlCommentIndicatorOnly"

    Dim c As Range, fullText As String, visibleText As String
    Dim added As Long, kept As Long, rowCount As Long

    Debug.Print "--------- Processing " & tgtCol.DataBodyRange.Rows.Count & " rows ---------"
    
    For Each c In tgtCol.DataBodyRange.Cells
        rowCount = rowCount + 1
        fullText = CStr(c.Value2)
        visibleText = SmartTruncate(fullText, maxLen)

        ' Wipe any old notes
        On Error Resume Next
        Dim hadComment As Boolean: hadComment = Not (c.Comment Is Nothing)
        If hadComment Then c.Comment.Delete
        On Error GoTo ErrHandler

        ' Always store something readable in the cell
        c.Value2 = visibleText

        ' Add the triangle only when we actually truncated
        If visibleText <> fullText Then
            c.AddComment Text:=fullText  ' Legacy NOTE → red triangle
            c.Comment.Visible = False
            added = added + 1
            If DebugMode Then Debug.Print "Row:" & c.Row, "Len:" & Len(fullText), "→ TRIANGLE", "Old comment:" & hadComment
        Else
            kept = kept + 1
            If DebugMode Then Debug.Print "Row:" & c.Row, "Len:" & Len(fullText), "(no change)", "Old comment:" & hadComment
        End If
        
        ' Progress indicator every 100 rows for large datasets
        If rowCount Mod 100 = 0 Then Debug.Print PROC & " - Processed " & rowCount & " rows..."
    Next c

    ' Final summary
    Debug.Print String(60, "-")
    Debug.Print PROC & " SUMMARY → Added:" & added & "  Unchanged:" & kept & "  Total rows:" & rowCount
    Debug.Print String(60, "-")
    
    ' Log results
    LogEvt PROC, lgINFO, "Added red-triangle notes to " & added & " rows. (Unchanged: " & kept & ", Total: " & rowCount & ")"
    Exit Sub

ErrHandler:
    Debug.Print PROC & " ERROR " & Err.Number & " – " & Err.Description
    LogEvt PROC, lgERROR, "Short-name / hover-note failed: " & Err.Description
End Sub

'-----------------------------------------------------------------
' Intelligently truncates text with an ellipsis, keeping words when possible.
'-----------------------------------------------------------------
Private Function SmartTruncate(txt As String, limit As Long) As String
    If Len(txt) <= limit Then
        SmartTruncate = txt
    Else
        Dim cut As Long: cut = InStrRev(Left$(txt, limit - 1), " ")
        If cut < 20 Then cut = limit - 1        'no space or too early? hard cut
        SmartTruncate = Left$(txt, cut) & "…"
    End If
End Function

Private Sub ReorganizeColumns(tbl As ListObject)
    ' Purpose: Moves columns to a predefined order.
    Const PROC_NAME As String = "mod_Format.ReorganizeColumns"
    Dim desiredOrder As Variant, currentPos As Long, targetPos As Long, colName As Variant, lc As ListColumn
    ' --- Define Desired Order (Consider moving to mod_Config) ---
    ' Updated order based on user feedback 2025-04-30 (Task: Fix formatting issues)
    ' Moved City and State to the end
    desiredOrder = Array( _
        "K_Number", "DecisionDate", "Applicant", "DeviceName", "Contact", _
        "CompanyRecap", "Score_Percent", "Category", "FDA_Link", _
        "Final_Score", "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "Country", "Statement", _
        "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", _
        "City", "State" _
    )
    ' Note: Any columns *not* listed here will end up even further to the right.

    On Error GoTo ReorgError
    LogEvt PROC_NAME, lgDETAIL, "Reorganizing columns...", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start reorganizing columns", "Table=" & tbl.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Entered ReorganizeColumns", tbl.Name ' <<< STANDALONE DEBUG

    Application.ScreenUpdating = False ' Speed up column moves

    targetPos = 1
    For Each colName In desiredOrder
        On Error Resume Next ' Check if column exists
        Set lc = tbl.ListColumns(colName)
        On Error GoTo ReorgError ' Restore handler

        If Not lc Is Nothing Then
            currentPos = lc.Index
            LogEvt PROC_NAME, lgDETAIL, "Processing column '" & colName & "'. Current Index=" & currentPos & ", Target Index=" & targetPos
            mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Processing column", "Col=" & colName & ", Current=" & currentPos & ", Target=" & targetPos
            StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Processing column", CStr(colName) & " | Current=" & currentPos & ", Target=" & targetPos ' <<< STANDALONE DEBUG
            If currentPos <> targetPos Then
                LogEvt PROC_NAME, lgDETAIL, "Attempting to move '" & colName & "' from " & currentPos & " to " & targetPos
                StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Attempting move", CStr(colName) & " | From=" & currentPos & ", To=" & targetPos ' <<< STANDALONE DEBUG
                lc.Range.EntireColumn.Cut
                tbl.HeaderRowRange.Parent.Columns(targetPos).Insert Shift:=xlToRight
                Application.CutCopyMode = False ' Clear clipboard
                LogEvt PROC_NAME, lgDETAIL, "Moved column '" & colName & "' from " & currentPos & " to " & targetPos, "Table=" & tbl.Name
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Moved column", "Table=" & tbl.Name & ", Col=" & colName & ", From=" & currentPos & ", To=" & targetPos
                StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Moved column", CStr(colName) ' <<< STANDALONE DEBUG
            End If
            targetPos = targetPos + 1
        Else
            LogEvt PROC_NAME, lgWARN, "Column '" & colName & "' not found for reorganization.", "Table=" & tbl.Name
            mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Column not found for reorg", "Table=" & tbl.Name & ", Col=" & colName
            StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "Column NOT FOUND for reorg", CStr(colName) ' <<< STANDALONE DEBUG
        End If
        Set lc = Nothing
    Next colName

    Application.ScreenUpdating = True
    LogEvt PROC_NAME, lgDETAIL, "Column reorganization complete.", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Column reorganization complete", "Table=" & tbl.Name
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ReorganizeColumns complete", tbl.Name ' <<< STANDALONE DEBUG
    Exit Sub

ReorgError:
    Application.ScreenUpdating = True ' Ensure screen updating is back on
    LogEvt PROC_NAME, lgERROR, "Error reorganizing columns: " & Err.Description, "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error reorganizing columns", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    StandaloneDebug.DebugLog "mod_Format", PROC_NAME, "ERROR in ReorganizeColumns", "Err=" & Err.Number & " - " & Err.Description ' <<< STANDALONE DEBUG
    MsgBox "An error occurred while reorganizing columns: " & Err.Description, vbExclamation, "Column Reorganization Error"
End Sub

Private Sub SortDataTable(tbl As ListObject)
    ' Purpose: Sorts the table by the primary sort key(s).
    Const PROC_NAME As String = "mod_Format.SortDataTable"
    Dim sortCol As Range
    On Error GoTo SortError
    LogEvt PROC_NAME, lgDETAIL, "Sorting data table...", "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start sorting table", "Table=" & tbl.Name

    ' --- Define Sort Key (Consider moving to mod_Config) ---
    Const SORT_COLUMN_NAME As String = "Final_Score"
    Dim SORT_ORDER As XlSortOrder ' Cannot use typed Const inside Sub
    SORT_ORDER = xlDescending

    On Error Resume Next ' Check if sort column exists
    Set sortCol = tbl.ListColumns(SORT_COLUMN_NAME).Range
    On Error GoTo SortError ' Restore handler

    If sortCol Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Sort column '" & SORT_COLUMN_NAME & "' not found. Skipping sort.", "Table=" & tbl.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Sort column not found", "Table=" & tbl.Name & ", Col=" & SORT_COLUMN_NAME
        Exit Sub
    End If

    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortCol, SortOn:=xlSortOnValues, Order:=SORT_ORDER, DataOption:=xlSortNormal
        ' Add secondary sort keys if needed:
        ' .SortFields.Add Key:=tbl.ListColumns("SecondaryColumn").Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    LogEvt PROC_NAME, lgDETAIL, "Data table sorted.", "Table=" & tbl.Name & ", Key=" & SORT_COLUMN_NAME
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Table sorted", "Table=" & tbl.Name & ", Key=" & SORT_COLUMN_NAME
    Exit Sub

SortError:
    LogEvt PROC_NAME, lgERROR, "Error sorting table: " & Err.Description, "Table=" & tbl.Name
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error sorting table", "Table=" & tbl.Name & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "An error occurred while sorting the table: " & Err.Description, vbExclamation, "Sort Error"
End Sub

Private Sub FreezeHeaderAndFirstColumns(ws As Worksheet)
    ' Purpose: Freezes the header row and the first few columns for better navigation.
    Const PROC_NAME As String = "mod_Format.FreezeHeaderAndFirstColumns"
    ' Const COLUMNS_TO_FREEZE As Long = 3 ' No longer used
    On Error Resume Next ' Be lenient with UI operations
    LogEvt PROC_NAME, lgDETAIL, "Freezing panes (DISABLED)...", "Sheet=" & ws.Name
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start freezing panes (DISABLED)", "Sheet=" & ws.Name

    ' --- Ensure panes are unfrozen ---
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False

    ' --- Original Freezing Logic (Commented Out) ---
    ' ws.Activate ' Must activate sheet to set freeze panes
    ' ActiveWindow.FreezePanes = False ' Unfreeze first
    ' With ws.Cells(2, COLUMNS_TO_FREEZE + 1) ' Cell below header, right of columns to freeze
    '     .Activate ' Select the cell to set the freeze boundary
    '     ActiveWindow.FreezePanes = True
    ' End With
    ' ws.Range("A1").Activate ' Select A1 after freezing
    ' --- End Original Logic ---

    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgWARN, "Error unfreezing panes: " & Err.Description, "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Error unfreezing panes", "Sheet=" & ws.Name & ", Err=" & Err.Description
        Err.Clear
    Else
        LogEvt PROC_NAME, lgDETAIL, "Panes unfrozen (or already unfrozen).", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Panes unfrozen", "Sheet=" & ws.Name
    End If
    On Error GoTo 0 ' Restore default
End Sub


' --- Helper function to check column existence (Added for logging) ---
Private Function ColumnExists(tbl As ListObject, colName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(colName)
    ColumnExists = (Err.Number = 0)
    On Error GoTo 0
    Set lc = Nothing
End Function


' ==========================================================================
' ===                  PUBLIC UTILITY FUNCTIONS                        ===
' ==========================================================================

Public Sub RestoreTableFormatting()
    ' Purpose: Manually restores all standard formatting to the main data table.
    Const PROC_NAME As String = "mod_Format.RestoreTableFormatting"
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim wsData As Worksheet ' Needed for ApplyAll call

    On Error GoTo RestoreError

    ' --- Get the main data sheet and table ---
    On Error Resume Next ' Handle sheet/table not found
    Set wsData = ThisWorkbook.Sheets(mod_Config.DATA_SHEET_NAME) ' Use config constant (Corrected name)
    If wsData Is Nothing Then
        MsgBox "Could not find the main data sheet: '" & mod_Config.DATA_SHEET_NAME & "'", vbCritical, "Restore Formatting Error" ' Corrected name
        Exit Sub
    End If
    Set tbl = wsData.ListObjects(1) ' Assumes first table on the sheet
    On Error GoTo RestoreError ' Restore handler

    If Not tbl Is Nothing Then
        LogEvt PROC_NAME, lgINFO, "Starting manual formatting restoration...", "Table=" & tbl.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start manual formatting restoration", "Table=" & tbl.Name

        ' --- Apply standard formatting sequence ---
        ' Note: ApplyAll handles its own internal logging and error handling
        If ApplyAll(tbl, wsData) Then
            LogEvt PROC_NAME, lgINFO, "Table formatting restoration completed successfully.", "Table=" & tbl.Name
            mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Manual formatting restoration complete", "Table=" & tbl.Name
            MsgBox "Table formatting has been restored successfully.", vbInformation, "Restore Complete"
        Else
            ' ApplyAll already showed an error message, just log it here
            LogEvt PROC_NAME, lgWARN, "Manual formatting restoration failed. See previous logs/messages.", "Table=" & tbl.Name
            mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Manual formatting restoration failed", "Table=" & tbl.Name
            ' No additional MsgBox needed here as ApplyAll handles it
        End If
    Else
        LogEvt PROC_NAME, lgERROR, "Could not find the main table on sheet '" & wsData.Name & "' to restore formatting."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Could not find table to restore", "Sheet=" & wsData.Name
        MsgBox "Could not find the main table on sheet '" & wsData.Name & "' to restore formatting.", vbExclamation, "Restore Formatting Error"
    End If

    Exit Sub

RestoreError:
    LogEvt PROC_NAME, lgERROR, "An unexpected error occurred during formatting restoration: " & Err.Description
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Unexpected error during restoration", "Err=" & Err.Number & " - " & Err.Description
    MsgBox "An unexpected error occurred during formatting restoration: " & vbCrLf & Err.Description, vbCritical, "Restore Formatting Error"
End Sub

' --- Helper function to get column index (Added for logging) ---
Private Function GetColumnIndex(tbl As ListObject, colName As String) As Long
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(colName)
    If Err.Number = 0 Then
        GetColumnIndex = lc.Index
    Else
        GetColumnIndex = 0 ' Return 0 if not found
    End If
    On Error GoTo 0
    Set lc = Nothing
End Function

' ==========================================================================
' ===                  TEST/DIAGNOSTIC FUNCTIONS                       ===
' ==========================================================================

Public Sub Test_DeviceNameTriangles()
    ' Purpose: Test function to quickly verify red triangle functionality
    ' without running the entire formatting process
    On Error GoTo TestError
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Get active sheet and its first table
    Set ws = ActiveSheet
    If ws.ListObjects.Count < 1 Then
        MsgBox "No table found on active sheet. Please select a sheet with a table.", vbExclamation, "Test Failed"
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(1)
    
    Debug.Print "================================================="
    Debug.Print "TESTING RED TRIANGLE FUNCTIONALITY"
    Debug.Print "Table: " & tbl.Name & " on sheet: " & ws.Name
    Debug.Print "================================================="
    
    ' Call with debug mode ON
    Debug.Print "Calling CreateShortNamesAndComments with default settings..."
    CreateShortNamesAndComments lo:=tbl, DebugMode:=True
    
    MsgBox "Test complete! Check the Immediate window (Ctrl+G) for detailed results." & vbCrLf & vbCrLf & _
           "If you don't see any red triangles, try these troubleshooting steps:" & vbCrLf & _
           "1. Make sure your DeviceName column has entries longer than 40 characters" & vbCrLf & _
           "2. Check Excel settings: File > Options > Advanced > Display > " & vbCrLf & _
           "   For cells with comments, show: > 'Indicators only (comments on hover)'", _
           vbInformation, "Red Triangle Test"
    Exit Sub
    
TestError:
    Debug.Print "ERROR in Test_DeviceNameTriangles: " & Err.Number & " - " & Err.Description
    MsgBox "Test failed: " & Err.Description, vbCritical, "Test Error"
End Sub

Public Sub Test_DeviceNameTriangles_ForceAll()
    ' Purpose: Force triangles on ALL device names by using a very small maxLen
    ' This is useful to verify the triangle functionality works at all
    On Error GoTo TestError
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Get active sheet and its first table
    Set ws = ActiveSheet
    If ws.ListObjects.Count < 1 Then
        MsgBox "No table found on active sheet. Please select a sheet with a table.", vbExclamation, "Test Failed"
        Exit Sub
    End If
    
    Set tbl = ws.ListObjects(1)
    
    Debug.Print "================================================="
    Debug.Print "FORCING RED TRIANGLES ON ALL DEVICE NAMES"
    Debug.Print "Table: " & tbl.Name & " on sheet: " & ws.Name
    Debug.Print "Setting maxLen=10 to force triangles on most entries"
    Debug.Print "================================================="
    
    ' Call with debug mode ON and small maxLen to force triangles
    Debug.Print "Calling CreateShortNamesAndComments with maxLen=10..."
    CreateShortNamesAndComments lo:=tbl, maxLen:=10, DebugMode:=True
    
    MsgBox "Test complete! Check the Immediate window (Ctrl+G) for detailed results." & vbCrLf & vbCrLf & _
           "This test forced triangles on most device names by limiting to 10 characters." & vbCrLf & _
           "If you still don't see any red triangles, check Excel settings:" & vbCrLf & _
           "File > Options > Advanced > Display > " & vbCrLf & _
           "For cells with comments, show: > 'Indicators only (comments on hover)'", _
           vbInformation, "Red Triangle Test (Forced)"
    Exit Sub
    
TestError:
    Debug.Print "ERROR in Test_DeviceNameTriangles_ForceAll: " & Err.Number & " - " & Err.Description
    MsgBox "Test failed: " & Err.Description, vbCritical, "Test Error"
End Sub
