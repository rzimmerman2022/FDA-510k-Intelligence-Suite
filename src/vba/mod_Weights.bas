' ==========================================================================
' Module      : mod_Weights
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module manages the loading and in-memory storage of
'               scoring weights and keyword lists used by the scoring engine
'               (mod_Score). It reads data from predefined Excel tables
'               (e.g., "tblACWeights", "tblKeywords") located on the worksheet
'               specified by WEIGHTS_SHEET_NAME in mod_Config. The loaded data
'               (weights into Dictionaries, keywords into Collections) is stored
'               in private module-level variables. Public accessor functions
'               provide read-only access to this cached data for other modules.
'
' Key Functions:
'               - LoadAll: Orchestrates the loading of all defined weight/keyword
'                 tables from the specified worksheet. Handles errors and logs
'                 success/failure for each table. Checks for critical load failures.
'               - GetACWeights, GetSTWeights, GetPCWeights: Public accessors
'                 returning the loaded weight dictionaries.
'               - GetHighValueKeywords, GetNFCosmeticKeywords, etc.: Public
'                 accessors returning the loaded keyword collections.
'
' Private Helpers:
'               - LoadTableToDict: Loads data from a 2-column table into a
'                 Scripting.Dictionary object.
'               - LoadTableToList: Loads data from the first column of a table
'                 into a Collection object.
'
' Dependencies: - mod_Logger: For logging loading progress and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - Assumes specific table names exist on the weights sheet.
'               - Requires Scripting.Dictionary object.
'               - Requires Collection object.
'               - Requires System.Collections.ArrayList object (in CheckKeywords).
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' 2025-04-30  Cline (AI)      - Qualified all TraceEvt calls with mod_DebugTraceHelpers.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit

' Module-level variables to hold the loaded data (kept Private)
Private dictACWeights As Object       ' Scripting.Dictionary: Key=AC Code, Value=Weight
Private dictSTWeights As Object       ' Scripting.Dictionary: Key=SubmType Code, Value=Weight
Private dictPCWeights As Object       ' Scripting.Dictionary: Key=PC Code, Value=Weight
Private highValKeywordsList As Collection ' Collection of high-value keywords (Strings)
Private nfCosmeticKeywordsList As Collection
Private nfDiagnosticKeywordsList As Collection
Private therapeuticKeywordsList As Collection

' Public entry point to load all tables
Public Function LoadAll(wsWeights As Worksheet) As Boolean
    ' Purpose: Loads all weight and keyword tables from the specified sheet into memory.
    Const PROC_NAME As String = "mod_Weights.LoadAll" ' Updated PROC_NAME
    Dim success As Boolean: success = True ' Assume success unless critical load fails
    On Error GoTo LoadErrorHandler ' General handler for non-critical table load issues

    LogEvt PROC_NAME, lgINFO, "Attempting to load weights and keywords from sheet: " & wsWeights.Name
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start loading", "Sheet='" & wsWeights.Name & "'"

    ' Load each table, log/trace success or failure
    Set dictACWeights = LoadTableToDict(wsWeights, "tblACWeights")
    mod_DebugTraceHelpers.TraceEvt IIf(dictACWeights Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded AC Weights", IIf(dictACWeights Is Nothing, "FAILED", "Count=" & dictACWeights.Count)
    Set dictSTWeights = LoadTableToDict(wsWeights, "tblSTWeights")
    mod_DebugTraceHelpers.TraceEvt IIf(dictSTWeights Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded ST Weights", IIf(dictSTWeights Is Nothing, "FAILED", "Count=" & dictSTWeights.Count)
    Set dictPCWeights = LoadTableToDict(wsWeights, "tblPCWeights") ' Optional
    mod_DebugTraceHelpers.TraceEvt IIf(dictPCWeights Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded PC Weights (Optional)", IIf(dictPCWeights Is Nothing, "FAILED/MISSING", "Count=" & dictPCWeights.Count)
    Set highValKeywordsList = LoadTableToList(wsWeights, "tblKeywords")
    mod_DebugTraceHelpers.TraceEvt IIf(highValKeywordsList Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded HighVal Keywords", IIf(highValKeywordsList Is Nothing, "FAILED", "Count=" & highValKeywordsList.Count)
    Set nfCosmeticKeywordsList = LoadTableToList(wsWeights, "tblNFCosmeticKeywords") ' Optional
    mod_DebugTraceHelpers.TraceEvt IIf(nfCosmeticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Cosmetic Keywords (Optional)", IIf(nfCosmeticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & nfCosmeticKeywordsList.Count)
    Set nfDiagnosticKeywordsList = LoadTableToList(wsWeights, "tblNFDiagnosticKeywords") ' Optional
    mod_DebugTraceHelpers.TraceEvt IIf(nfDiagnosticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Diagnostic Keywords (Optional)", IIf(nfDiagnosticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & nfDiagnosticKeywordsList.Count)
    Set therapeuticKeywordsList = LoadTableToList(wsWeights, "tblTherapeuticKeywords") ' Optional
    mod_DebugTraceHelpers.TraceEvt IIf(therapeuticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Therapeutic Keywords (Optional)", IIf(therapeuticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & therapeuticKeywordsList.Count)

    ' --- Critical Check: Ensure essential tables were loaded ---
    If dictACWeights Is Nothing Or dictSTWeights Is Nothing Or highValKeywordsList Is Nothing Then
         LogEvt PROC_NAME, lgERROR, "Critical failure: Could not load AC/ST weights or HighValue Keywords."
         mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "CRITICAL FAILURE: Missing essential Weights/Keywords", "AC=" & IIf(dictACWeights Is Nothing, "FAIL", "OK") & ", ST=" & IIf(dictSTWeights Is Nothing, "FAIL", "OK") & ", KW=" & IIf(highValKeywordsList Is Nothing, "FAIL", "OK")
        GoTo LoadErrorCritical ' Jump to specific critical error handling
    End If

    ' --- Standard Logging (Summary for RunLog) ---
    LogEvt PROC_NAME, IIf(dictACWeights.Count = 0, lgWARN, lgDETAIL), "Loaded " & dictACWeights.Count & " AC Weights."
    LogEvt PROC_NAME, IIf(dictSTWeights.Count = 0, lgWARN, lgDETAIL), "Loaded " & dictSTWeights.Count & " ST Weights."
    LogEvt PROC_NAME, IIf(dictPCWeights Is Nothing Or dictPCWeights.Count = 0, lgINFO, lgDETAIL), "Loaded " & IIf(dictPCWeights Is Nothing, 0, dictPCWeights.Count) & " PC Weights (Optional)."
    LogEvt PROC_NAME, IIf(highValKeywordsList.Count = 0, lgWARN, lgDETAIL), "Loaded " & highValKeywordsList.Count & " HighVal Keywords."
    LogEvt PROC_NAME, IIf(nfCosmeticKeywordsList Is Nothing Or nfCosmeticKeywordsList.Count = 0, lgINFO, lgDETAIL), "Loaded " & IIf(nfCosmeticKeywordsList Is Nothing, 0, nfCosmeticKeywordsList.Count) & " Cosmetic Keywords (Optional)."
    LogEvt PROC_NAME, IIf(nfDiagnosticKeywordsList Is Nothing Or nfDiagnosticKeywordsList.Count = 0, lgINFO, lgDETAIL), "Loaded " & IIf(nfDiagnosticKeywordsList Is Nothing, 0, nfDiagnosticKeywordsList.Count) & " Diagnostic Keywords (Optional)."
    LogEvt PROC_NAME, IIf(therapeuticKeywordsList Is Nothing Or therapeuticKeywordsList.Count = 0, lgINFO, lgDETAIL), "Loaded " & IIf(therapeuticKeywordsList Is Nothing, 0, therapeuticKeywordsList.Count) & " Therapeutic Keywords (Optional)."

    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Loading complete", "Success=True"
    LoadAll = True ' Indicate success
    Exit Function

LoadErrorHandler: ' Handles non-critical errors (e.g., optional table missing)
    Dim errDesc As String: errDesc = Err.Description
     LogEvt PROC_NAME, lgWARN, "Non-critical error loading one or more weight/keyword tables: " & errDesc & ". Defaults may be used.", "Sheet=" & wsWeights.Name
     mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Non-critical load error occurred", "Sheet='" & wsWeights.Name & "', Err=" & Err.Number & " - " & errDesc
    ' Don't MsgBox here, allow process to continue with defaults if possible
    ' Ensure objects are initialized even if loading failed, to prevent later errors
    If dictACWeights Is Nothing Then Set dictACWeights = CreateObject("Scripting.Dictionary"): dictACWeights.CompareMode = vbTextCompare
    If dictSTWeights Is Nothing Then Set dictSTWeights = CreateObject("Scripting.Dictionary"): dictSTWeights.CompareMode = vbTextCompare
    If dictPCWeights Is Nothing Then Set dictPCWeights = CreateObject("Scripting.Dictionary"): dictPCWeights.CompareMode = vbTextCompare
    If highValKeywordsList Is Nothing Then Set highValKeywordsList = New Collection
    If nfCosmeticKeywordsList Is Nothing Then Set nfCosmeticKeywordsList = New Collection
    If nfDiagnosticKeywordsList Is Nothing Then Set nfDiagnosticKeywordsList = New Collection
    If therapeuticKeywordsList Is Nothing Then Set therapeuticKeywordsList = New Collection
    ' Resume Next ' Resume execution after handling non-critical error (implicit with On Error GoTo)
     LoadAll = True ' Still return True for non-critical errors, allowing defaults
     Exit Function ' Need to explicitly exit after handling non-critical error

LoadErrorCritical: ' Handles failure to load essential tables
    MsgBox "Critical Error: Could not load essential AC/ST weights or HighValue Keywords from sheet '" & wsWeights.Name & "'. Processing cannot continue.", vbCritical, "Load Failure"
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Exiting due to critical load failure."
    ' Clean up any potentially partially loaded objects
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    LoadAll = False ' Indicate critical failure
End Function

' Keep these helpers Private within this module
Private Function LoadTableToDict(ws As Worksheet, tableName As String) As Object ' Scripting.Dictionary or Nothing
    ' Purpose: Loads a 2-column table into a Dictionary. Returns Nothing on error.
    Dim dict As Object ' Late bound Scripting.Dictionary
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, key As String, val As Variant
    Const PROC_NAME As String = "mod_Weights.LoadTableToDict" ' Updated PROC_NAME

    On Error GoTo LoadDictError

    Set tbl = ws.ListObjects(tableName) ' This will error if table doesn't exist

    If tbl.ListRows.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "Table '" & tableName & "' is empty. Returning empty dictionary.", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Table empty", "Table=" & tableName
        Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare ' Return empty dict
        Set LoadTableToDict = dict
        Exit Function
    End If

    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then ' Should not happen if rows > 0, but check
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has rows but no DataBodyRange? Returning Nothing.", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "No DataBodyRange despite rows>0", "Table=" & tableName
        Set LoadTableToDict = Nothing ' Indicate failure
        Exit Function
    End If

    If dataRange.Columns.Count < 2 Then
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has less than 2 columns. Cannot create dictionary. Returning Nothing.", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Table has < 2 columns", "Table=" & tableName
        Set LoadTableToDict = Nothing ' Indicate failure
        Exit Function
    End If

    ' Read data into array
    dataArr = dataRange.Value2 ' Use Value2 for raw data

    ' Create and populate dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If IsArray(dataArr) Then
        If UBound(dataArr, 2) >= 2 Then ' Ensure it's a 2D array with at least 2 columns
            For i = 1 To UBound(dataArr, 1)
                key = Trim(CStr(dataArr(i, 1))) ' Key from first column
                val = dataArr(i, 2)             ' Value from second column
                If Len(key) > 0 Then
                    dict(key) = val ' Add or overwrite
                End If
            Next i
        Else ' Handle case where single row might return 1D array
            If tbl.ListRows.Count = 1 And UBound(dataArr) >= 2 Then ' Check UBound for 1D array elements
                 key = Trim(CStr(dataArr(1)))
                 val = dataArr(2)
                 If Len(key) > 0 Then dict(key) = val
            Else
                  LogEvt PROC_NAME, lgWARN, "Unexpected array structure from table '" & tableName & "'. Cannot populate dictionary.", "Sheet=" & ws.Name
                  mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Unexpected array structure", "Table=" & tableName
                  Set LoadTableToDict = Nothing ' Indicate failure
                  Exit Function
            End If
        End If
    ElseIf Not IsEmpty(dataArr) And tbl.ListRows.Count = 1 And tbl.ListColumns.Count >= 2 Then
         ' Handle single row, non-array result (less common with Value2)
         key = Trim(CStr(tbl.DataBodyRange.Cells(1, 1).Value2))
         val = tbl.DataBodyRange.Cells(1, 2).Value2
         If Len(key) > 0 Then dict(key) = val
    End If

    Set LoadTableToDict = dict ' Return populated dictionary
    Exit Function

LoadDictError:
    LogEvt PROC_NAME, lgWARN, "Error loading table '" & tableName & "' to Dict: " & Err.Description & ". Returning Nothing.", "Sheet=" & ws.Name
    mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Error loading table to Dict", "Table=" & tableName & ", Err=" & Err.Number & " - " & Err.Description
    Set LoadTableToDict = Nothing ' Return Nothing on error
    ' No Resume here, exit handled by returning Nothing
End Function

Private Function LoadTableToList(ws As Worksheet, tableName As String) As Collection ' Returns Collection or Nothing
    ' Purpose: Loads the first column of a table into a Collection. Returns Nothing on error.
    Dim coll As Collection ' Use New Collection if early bound, otherwise create later
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, item As String
    Const PROC_NAME As String = "mod_Weights.LoadTableToList" ' Updated PROC_NAME

    On Error GoTo LoadListError

    Set tbl = ws.ListObjects(tableName) ' Errors if table doesn't exist

    If tbl.ListRows.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "Table '" & tableName & "' is empty. Returning empty collection.", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Table empty", "Table=" & tableName
        Set coll = New Collection ' Return new empty collection
        Set LoadTableToList = coll
        Exit Function
    End If

    Set dataRange = tbl.ListColumns(1).DataBodyRange ' Get first column data
    If dataRange Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has rows but no DataBodyRange in first column? Returning Nothing.", "Sheet=" & ws.Name
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "No DataBodyRange in Col1", "Table=" & tableName
        Set LoadTableToList = Nothing
        Exit Function
    End If

    dataArr = dataRange.Value2 ' Read data

    Set coll = New Collection ' Create collection object

    If IsArray(dataArr) Then
        For i = 1 To UBound(dataArr, 1) ' Assumes vertical array from column read
            item = Trim(CStr(dataArr(i, 1)))
            If Len(item) > 0 Then
                On Error Resume Next ' Ignore error if item already exists (use collection as unique list)
                coll.Add item, item ' Add unique items using item itself as key
                On Error GoTo LoadListError ' Restore handler
            End If
        Next i
    ElseIf Not IsEmpty(dataArr) Then ' Handle single row data
        item = Trim(CStr(dataArr))
        If Len(item) > 0 Then
             On Error Resume Next ' Ignore error if adding single item fails (e.g., key exists if called strangely)
             coll.Add item, item
             On Error GoTo LoadListError ' Restore handler
        End If
    End If

    Set LoadTableToList = coll ' Return populated collection
    Exit Function

LoadListError:
    LogEvt PROC_NAME, lgWARN, "Error loading table '" & tableName & "' to List: " & Err.Description & ". Returning Nothing.", "Sheet=" & ws.Name
    mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Error loading table to List", "Table=" & tableName & ", Err=" & Err.Number & " - " & Err.Description
    Set LoadTableToList = Nothing ' Return Nothing on error
End Function

' --- Public Accessors for Loaded Data ---
' These functions provide controlled access to the private module-level dictionaries/collections.

Public Function GetACWeights() As Object ' Scripting.Dictionary
    ' Returns the loaded AC Weights dictionary. Returns empty dictionary if not loaded.
    If dictACWeights Is Nothing Then Set dictACWeights = CreateObject("Scripting.Dictionary"): dictACWeights.CompareMode = vbTextCompare
    Set GetACWeights = dictACWeights
End Function

Public Function GetSTWeights() As Object ' Scripting.Dictionary
    ' Returns the loaded ST Weights dictionary. Returns empty dictionary if not loaded.
    If dictSTWeights Is Nothing Then Set dictSTWeights = CreateObject("Scripting.Dictionary"): dictSTWeights.CompareMode = vbTextCompare
    Set GetSTWeights = dictSTWeights
End Function

Public Function GetPCWeights() As Object ' Scripting.Dictionary
    ' Returns the loaded PC Weights dictionary. Returns empty dictionary if not loaded.
    If dictPCWeights Is Nothing Then Set dictPCWeights = CreateObject("Scripting.Dictionary"): dictPCWeights.CompareMode = vbTextCompare
    Set GetPCWeights = dictPCWeights
End Function

Public Function GetHighValueKeywords() As Collection
    ' Returns the loaded High Value Keywords collection. Returns empty collection if not loaded.
    If highValKeywordsList Is Nothing Then Set highValKeywordsList = New Collection
    Set GetHighValueKeywords = highValKeywordsList
End Function

Public Function GetNFCosmeticKeywords() As Collection
    ' Returns the loaded Cosmetic Keywords collection. Returns empty collection if not loaded.
    If nfCosmeticKeywordsList Is Nothing Then Set nfCosmeticKeywordsList = New Collection
    Set GetNFCosmeticKeywords = nfCosmeticKeywordsList
End Function

Public Function GetNFDiagnosticKeywords() As Collection
    ' Returns the loaded Diagnostic Keywords collection. Returns empty collection if not loaded.
    If nfDiagnosticKeywordsList Is Nothing Then Set nfDiagnosticKeywordsList = New Collection
    Set GetNFDiagnosticKeywords = nfDiagnosticKeywordsList
End Function

Public Function GetTherapeuticKeywords() As Collection
    ' Returns the loaded Therapeutic Keywords collection. Returns empty collection if not loaded.
    If therapeuticKeywordsList Is Nothing Then Set therapeuticKeywordsList = New Collection
    Set GetTherapeuticKeywords = therapeuticKeywordsList
End Function
