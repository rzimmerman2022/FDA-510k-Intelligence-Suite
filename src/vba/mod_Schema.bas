' =========  mod_Schema.bas  =========
' Purpose: Handles table schema operations, including column index mapping,
'          safe data retrieval by column name, and header validation.
' Key APIs exposed: GetColumnIndices, ColumnExistsInMap, SafeGetString, SafeGetVariant, etc.
' Maintainer: [Your Name/Team]
' Dependencies: mod_Logger, mod_DebugTraceHelpers
' =====================================
Option Explicit
Attribute VB_Name = "mod_Schema"

Public Function GetColumnIndices(headerRange As Range) As Object ' Scripting.Dictionary or Nothing
    ' Purpose: Creates a dictionary mapping column header names (handling duplicates) to their 1-based index.
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive for lookups by base name
    Dim cell As Range, colNum As Long, missingCols As String, h As String
    Dim dupeCheckDict As Object: Set dupeCheckDict = CreateObject("Scripting.Dictionary")
    dupeCheckDict.CompareMode = vbBinaryCompare ' Case-sensitive check for exact duplicates
    Const PROC_NAME As String = "GetColumnIndices"
    colNum = 1

    Dim requiredBaseCols As Variant ' Base names of columns that MUST exist
    requiredBaseCols = Array("K_Number", "Applicant", "DeviceName", "DecisionDate", "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "Country", "Statement", "FDA_Link")
    ' Output columns are also required, added by AddScoreColumnsIfNeeded, listed here for check completeness
    Dim requiredOutputCols As Variant
    requiredOutputCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")

    TraceEvt lvlINFO, PROC_NAME, "Mapping header range", "Range=" & headerRange.Address(External:=True)

    For Each cell In headerRange.Cells
        h = Trim(cell.Value)
        If Len(h) > 0 Then
            Dim dictKey As String
            ' --- Handle potential duplicate header names ---
            If dupeCheckDict.Exists(h) Then
                ' Exact duplicate header found - create unique key using column index
                dictKey = h & "#" & colNum
                LogEvt PROC_NAME, lgWARN, "Duplicate header '" & h & "' detected. Using unique key: '" & dictKey & "' for column index " & colNum
                TraceEvt lvlWARN, PROC_NAME, "Duplicate header detected", "Header='" & h & "', UsingKey='" & dictKey & "', Index=" & colNum
                ' Increment count for original duplicate check dictionary (optional, for info)
                dupeCheckDict(h) = dupeCheckDict(h) + 1
            Else
                ' First time seeing this header name
                dictKey = h ' Use the plain name as the key
                dupeCheckDict.Add h, 1 ' Add to duplicate checker
            End If
            ' --- End Duplicate Handling ---

            ' Add to the main dictionary used for lookups (case-insensitive key)
            If Not dict.Exists(dictKey) Then
                dict.Add dictKey, colNum
            Else
                ' This case should be rare with the Name#Index logic, but log if it occurs
                LogEvt PROC_NAME, lgWARN, "Duplicate key '" & dictKey & "' encountered in main mapping dict for column index " & colNum & ". Check header processing logic."
                TraceEvt lvlWARN, PROC_NAME, "Duplicate key conflict in map", "Key='" & dictKey & "', Index=" & colNum
            End If
        Else
             LogEvt PROC_NAME, lgDETAIL, "Skipping blank header cell.", "Index=" & colNum
             TraceEvt lvlDET, PROC_NAME, "Skipping blank header cell", "Index=" & colNum
        End If
        colNum = colNum + 1
    Next cell
    Set dupeCheckDict = Nothing ' Clean up temp dictionary

    ' --- Check for missing REQUIRED columns (using base names) ---
    Dim allRequiredCols As Variant: allRequiredCols = ConcatArrays(requiredBaseCols, requiredOutputCols)
    Dim reqCol As Variant
    For Each reqCol In allRequiredCols
        If Not ColumnExistsInMap(dict, reqCol) Then
            missingCols = missingCols & vbCrLf & " - " & reqCol
        End If
    Next reqCol

    If Len(missingCols) > 0 Then
         LogEvt PROC_NAME, lgERROR, "Required columns missing in table header:" & Replace(missingCols, vbCrLf, ", ")
         TraceEvt lvlERROR, PROC_NAME, "Required columns missing", "Missing=" & Replace(missingCols, vbCrLf, ", ")
        MsgBox "Error: The following required columns were not found in sheet '" & headerRange.Parent.Name & "':" & missingCols & vbCrLf & "Please ensure Power Query output and VBA column additions match.", vbCritical, "Missing Columns"
        Set GetColumnIndices = Nothing ' Return Nothing on failure
    Else
         LogEvt PROC_NAME, lgINFO, "Column indices mapped successfully.", "MappedKeys=" & dict.Count
         TraceEvt lvlINFO, PROC_NAME, "Column mapping successful", "MappedKeys=" & dict.Count
        Set GetColumnIndices = dict ' Return the populated dictionary
    End If
End Function

Public Function ColumnExistsInMap(dict As Object, baseColName As String) As Boolean
    ' Helper to check if a base column name exists in the dictionary,
    ' either as a direct key or as part of a Name#Index key.
    ColumnExistsInMap = False
    If dict Is Nothing Then Exit Function
    If dict.Exists(baseColName) Then ' Check direct match first
        ColumnExistsInMap = True
        Exit Function
    End If
    ' Check for Name#Index format
    Dim itemKey As Variant
    For Each itemKey In dict.Keys
        If itemKey Like baseColName & "#*" Then
            ColumnExistsInMap = True
            Exit Function
        End If
    Next itemKey
End Function

Public Function ConcatArrays(arr1 As Variant, arr2 As Variant) As Variant
    ' Simple helper to concatenate two 1D arrays
    Dim tempArr() As Variant
    Dim i As Long, j As Long, size1 As Long, size2 As Long
    size1 = UBound(arr1) - LBound(arr1) + 1
    size2 = UBound(arr2) - LBound(arr2) + 1
    ReDim tempArr(1 To size1 + size2)
    j = 1
    For i = LBound(arr1) To UBound(arr1)
        tempArr(j) = arr1(i)
        j = j + 1
    Next i
    For i = LBound(arr2) To UBound(arr2)
        tempArr(j) = arr2(i)
        j = j + 1
    Next i
    ConcatArrays = tempArr
End Function

Public Function SafeGetString(arr As Variant, r As Long, ByVal cols As Object, baseColName As String) As String
    ' Purpose: Safely gets a string value from the data array using the column map, handling missing/duplicate columns.
    Dim colIdx As Long
    colIdx = SafeGetColIndex(cols, baseColName) ' Find the correct index

    If colIdx > 0 Then
        On Error Resume Next ' Handle error reading specific array element
        SafeGetString = Trim(CStr(arr(r, colIdx)))
        If Err.Number <> 0 Then SafeGetString = "": Err.Clear ' Return blank on error
        On Error GoTo 0
    Else
        SafeGetString = "" ' Return blank if column index not found
    End If
End Function

Public Function SafeGetVariant(arr As Variant, r As Long, ByVal cols As Object, baseColName As String) As Variant
    ' Purpose: Safely gets a variant value from the data array using the column map.
    Dim colIdx As Long
    colIdx = SafeGetColIndex(cols, baseColName)

    If colIdx > 0 Then
        On Error Resume Next ' Handle error reading specific array element
        SafeGetVariant = arr(r, colIdx)
        If Err.Number <> 0 Then SafeGetVariant = Null: Err.Clear ' Return Null on error (or CVErr?)
        On Error GoTo 0
    Else
        SafeGetVariant = Null ' Return Null if column index not found
    End If
End Function

Public Function SafeGetColIndex(colsDict As Object, baseColName As String) As Long
    ' Purpose: Finds the column index from the dictionary, trying base name first, then Name#Index format.
    SafeGetColIndex = 0 ' Default to 0 (not found)
    If colsDict Is Nothing Then Exit Function

    On Error Resume Next ' Ignore errors during dictionary lookup

    ' 1. Try direct lookup by base name
    SafeGetColIndex = CLng(colsDict(baseColName))
    If Err.Number = 0 And SafeGetColIndex > 0 Then Exit Function ' Found it directly
    Err.Clear

    ' 2. If not found, iterate keys to find Name#Index format
    Dim itemKey As Variant
    For Each itemKey In colsDict.Keys
        If itemKey Like baseColName & "#*" Then
            SafeGetColIndex = CLng(colsDict(itemKey))
            If Err.Number = 0 And SafeGetColIndex > 0 Then Exit Function ' Found it with #Index
            Err.Clear
        End If
    Next itemKey

    ' If we reach here, the column wasn't found by either method
    SafeGetColIndex = 0
    ' LogEvt "SafeGetColIndex", lgDETAIL, "Column base name not found in map.", "BaseName=" & baseColName ' Log only if truly not found - Can be noisy
    TraceEvt lvlDET, "SafeGetColIndex", "Column base name not found in map", "BaseName=" & baseColName
    On Error GoTo 0 ' Restore default error handling
End Function
