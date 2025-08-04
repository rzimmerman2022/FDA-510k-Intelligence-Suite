'--- Code for Module: mod_510k_Processor ---
Option Explicit

' ==========================================================================
' ==========================================================================
' ===                                                                    ===
' ===          FDA 510(k) LEAD SCORING PROCESSING MODULE                 ===
' ===                 (mod_510k_Processor)                               ===
' ===                                                                    ===
' ==========================================================================
' ==========================================================================
' --- Requires companion module: mod_Logger (Optional but Recommended) ---
' --- Assumes Excel Tables named: tblACWeights, tblSTWeights,          ---
' --- tblPCWeights, tblKeywords on sheet named "Weights"               ---
' --- Assumes Cache sheet named "CompanyCache" with headers            ---
' --- Assumes Data sheet named "CurrentMonthData" with PQ output table ---

' ==========================================================================
' ===                     CONFIGURATION CONSTANTS                        ===
' ==========================================================================
' --- Essential Configuration ---
' *** IMPORTANT: SET YOUR WINDOWS USERNAME FOR MAINTAINER FEATURES (e.g., OpenAI, DebugMode) ***
Private Const MAINTAINER_USERNAME As String = "YourWindowsUsername" ' <<< UPDATE THIS

' *** Double-check these names match your Excel objects ***
' Public Const PQ_CONNECTION_NAME As String = "pqGet510kData" ' <<< No longer needed/used by RefreshPowerQuery
Public Const DATA_SHEET_NAME As String = "CurrentMonthData"  ' Sheet where Power Query loads data (Public for ThisWorkbook)
Private Const WEIGHTS_SHEET_NAME As String = "Weights"        ' Sheet containing weight/keyword tables
Private Const CACHE_SHEET_NAME As String = "CompanyCache"      ' Sheet for persistent company recap cache
Private Const LOG_SHEET_NAME As String = "LogSheet"             ' Optional: Name for the log sheet (used by Logger module)
Private Const SHORT_NAME_MAX_LEN As Long = 75 ' Maximum length for shortened device names
Private Const SHORT_NAME_ELLIPSIS As String = "..." ' Text to append to shortened names

' *** Path to file containing ONLY your OpenAI API Key ***
' *** Uses %APPDATA% environment variable for user-specific location ***
Private Const API_KEY_FILE_PATH As String = "%APPDATA%\510k_Tool\openai_key.txt" ' <<< ENSURE THIS PATH IS CORRECT & FILE EXISTS

' --- Scoring Defaults & Parameters (Used if lookup fails or as base values) ---
' *** REVIEW AND CONFIRM THESE VALUES BASED ON YOUR SCORING MODEL ***
Private Const DEFAULT_AC_WEIGHT As Double = 0.2
Private Const DEFAULT_PC_WEIGHT As Double = 0.2
Private Const DEFAULT_ST_WEIGHT As Double = 0.6 ' Default to Traditional if SubmType not found
Private Const DEFAULT_PT_WEIGHT As Double = 0.5 ' Default if ProcTimeDays is invalid or <162
Private Const HIGH_KW_WEIGHT As Double = 0.85
Private Const LOW_KW_WEIGHT As Double = 0.2
Private Const US_GL_WEIGHT As Double = 0.6
Private Const OTHER_GL_WEIGHT As Double = 0.5
Private Const NF_COSMETIC As Double = -2#  ' Negative Factor for purely cosmetic devices (CONFIRM VALUE)
Private Const NF_DIAGNOSTIC As Double = -0.2 ' Negative Factor for purely diagnostic software (CONFIRM VALUE)
Private Const SYNERGY_BONUS As Double = 0.15 ' Bonus for specific AC + High KW match (CONFIRM VALUE/LOGIC)

' --- OpenAI Configuration (Optional) ---
Private Const OPENAI_API_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const OPENAI_MODEL As String = "gpt-3.5-turbo" ' Or "gpt-4o-mini" etc. - check pricing/availability
Private Const OPENAI_MAX_TOKENS As Long = 100 ' Limit response length
Private Const OPENAI_TIMEOUT_MS As Long = 60000 ' 60 seconds timeout for API call

' --- UI & Formatting ---
Public Const VERSION_INFO As String = "v1.1 - Refactored Refresh/Reorg" ' Simple version tracking (Public for Logger)
Private Const RECAP_MAX_LEN As Long = 32760 ' Max characters for cell / recap text to avoid overflow
Private Const DEFAULT_RECAP_TEXT = "Needs Research" ' Default text when recap is missing

' ==========================================================================
' ===               MODULE-LEVEL VARIABLES / OBJECTS                   ===
' ==========================================================================
' Dictionaries and Collections for efficient lookups during processing
Private dictACWeights As Object       ' Scripting.Dictionary: Key=AC Code, Value=Weight
Private dictSTWeights As Object       ' Scripting.Dictionary: Key=SubmType Code, Value=Weight
Private dictPCWeights As Object       ' Scripting.Dictionary: Key=PC Code, Value=Weight
Private highValKeywordsList As Collection ' Collection of high-value keywords (Strings)
Private nfCosmeticKeywordsList As Collection
Private nfDiagnosticKeywordsList As Collection
Private therapeuticKeywordsList As Collection
Private dictCache As Object           ' Scripting.Dictionary: Key=CompanyName, Value=RecapText (In-memory for current run)


' ==========================================================================
' ===                   MAIN ORCHESTRATION SUB                         ===
' ==========================================================================
Public Sub ProcessMonthly510k()
    ' Purpose: Main control routine orchestrating the entire 510(k) processing workflow.
    '          Called by Workbook_Open or a button. Handles setup, data refresh,
    '          parameter loading, scoring, formatting, caching, and archiving.

    ' --- Variable Declarations ---
    Dim wsData As Worksheet, wsWeights As Worksheet, wsCache As Worksheet, wsLog As Worksheet
    Dim startMonth As Date, targetMonthName As String, archiveSheetName As String
    Dim startTime As Double: startTime = Timer ' Start timing the process
    Dim tblData As ListObject ' Represents the main data table
    Dim recordCount As Long   ' Number of records to process
    Dim dataArr As Variant    ' Array to hold data for fast processing
    Dim i As Long             ' Loop counter
    Dim scoreResult As Variant ' Array holding results from Calculate510kScore
    Dim currentRecap As String ' Holds the company recap text
    Dim useOpenAI As Boolean   ' Flag indicating if OpenAI should be attempted
    Dim colIndices As Object   ' Dictionary mapping column names to indices
    Dim proceed As Boolean     ' Flag indicating if processing should run
    Dim mustArchive As Boolean ' Flag indicating if archiving is needed

    ' --- Error Handling Setup ---
    On Error GoTo ProcessErrorHandler

    ' --- Initial Setup & Screen Handling ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = "Initializing 510(k) processing..."

    ' --- Initialize Logging ---
    LogEvt "ProcessStart", lvlINFO, "ProcessMonthly510k Started", "Version=" & VERSION_INFO

    ' --- Get Worksheet Objects Safely ---
    If Not GetWorksheets(wsData, wsWeights, wsCache) Then GoTo CleanExit

    ' --- Determine Target Month & Check Guard Conditions ---
    startMonth = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
    targetMonthName = Format$(startMonth, "MMM-yyyy")
    archiveSheetName = targetMonthName
    mustArchive = Not SheetExists(archiveSheetName)
    proceed = mustArchive Or Day(Date) <= 5 Or IsMaintainerUser()

    LogEvt "ArchiveCheck", IIf(proceed, lvlINFO, lvlWARN), _
           "Guard conditions: Archive needed=" & mustArchive & _
           ", Day of month=" & Day(Date) & ", Is maintainer=" & IsMaintainerUser() & _
           ", Will proceed=" & proceed

    If Not proceed Then
        LogEvt "ProcessSkip", lvlINFO, "Processing skipped: Archive exists, not day 1-5, not maintainer."
        Application.StatusBar = "Month " & targetMonthName & " already archived. Refreshing current view only."
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
        If tblData Is Nothing Then
            LogEvt "Refresh", lvlERROR, "Data table not found on " & DATA_SHEET_NAME & " during skipped run check."
        Else
            If Not RefreshPowerQuery(tblData) Then
                LogEvt "Refresh", lvlERROR, "PQ Refresh failed during skipped run check."
            End If
        End If
        Set tblData = Nothing
        GoTo CleanExit
    End If
    Application.StatusBar = "Processing for month: " & targetMonthName

    ' --- Get Data Table & Check for Data ---
    If tblData Is Nothing Then
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
    End If
    If tblData Is Nothing Then
        LogEvt "DataTable", lvlERROR, "Data table not found on " & DATA_SHEET_NAME
        GoTo ProcessErrorHandler
    End If

    ' --- Refresh Power Query Data (using the table object) ---
    Application.StatusBar = "Refreshing FDA data from Power Query..."
    LogEvt "Refresh", lvlINFO, "Attempting PQ refresh for table: " & tblData.Name
    If Not RefreshPowerQuery(tblData) Then GoTo ProcessErrorHandler ' Stop on critical PQ error

    LogEvt "DataTable", lvlINFO, "Found data table: " & tblData.Name
    If tblData.ListRows.Count = 0 Then
        LogEvt "DataTable", lvlWARN, "No data returned by Power Query for " & targetMonthName & "."
        MsgBox "No data returned by Power Query for " & targetMonthName & ". Nothing to process.", vbInformation, "No Data"
        GoTo CleanExit
    End If
    recordCount = tblData.ListRows.Count
    LogEvt "DataTable", lvlINFO, "Table contains " & recordCount & " rows."

    ' --- Add/Verify Output Columns ---
    LogEvt "Columns", lvlINFO, "Checking/Adding scoring output columns..."
    If Not AddScoreColumnsIfNeeded(tblData) Then GoTo ProcessErrorHandler

    ' --- Map Column Headers to Indices ---
    Set colIndices = GetColumnIndices(tblData.HeaderRowRange)
    If colIndices Is Nothing Then GoTo ProcessErrorHandler

    ' --- Load Weights, Keywords, and Cache ---
    Application.StatusBar = "Loading scoring parameters and cache..."
    LogEvt "LoadParams", lvlINFO, "Loading weights, keywords, and cache..."
    If Not LoadWeightsAndKeywords(wsWeights) Then GoTo ProcessErrorHandler
    Call LoadCompanyCache(wsCache)

    ' --- Read Data into Array for Fast Processing ---
    Application.StatusBar = "Reading data into memory (" & recordCount & " rows)..."
    LogEvt "ReadData", lvlINFO, "Reading data into array..."
    If recordCount = 1 Then
        Dim singleRowData As Variant
        singleRowData = tblData.DataBodyRange.Value
        If IsArray(singleRowData) Then
             If UBound(singleRowData, 1) = 1 And UBound(singleRowData, 2) > 0 Then
                 dataArr = singleRowData
             Else
                 ReDim dataArr(1 To 1, 1 To tblData.ListColumns.Count)
                 Dim j As Long
                 Dim tempArr As Variant: tempArr = tblData.DataBodyRange.Value2
                 For j = 1 To tblData.ListColumns.Count
                     On Error Resume Next
                     dataArr(1, j) = tempArr(j)
                     On Error GoTo 0
                 Next j
                 LogEvt "ReadData", lvlDETAIL, "Manually created 2D array for single row."
             End If
        Else
             LogEvt "ReadData", lvlERROR, "Failed to read single row into an array."
             GoTo ProcessErrorHandler
        End If
    ElseIf recordCount > 1 Then
        dataArr = tblData.DataBodyRange.Value2
    Else
        LogEvt "ReadData", lvlWARN, "Attempted to read data array when recordCount is 0."
        GoTo CleanExit
    End If
    LogEvt "ReadData", lvlINFO, "Read " & recordCount & " records into array (Ensured 2D)."

    ' --- Main Processing Loop ---
    Application.StatusBar = "Calculating scores and fetching recaps (0% Complete)..."
    LogEvt "ScoreLoop", lvlINFO, "Starting main processing loop for " & recordCount & " records."
    useOpenAI = IsMaintainerUser()

    For i = 1 To recordCount
        scoreResult = Calculate510kScore(dataArr, i, colIndices)
        Dim companyName As String
        On Error Resume Next: companyName = Trim(CStr(dataArr(i, colIndices("Applicant")))): On Error GoTo ProcessErrorHandler
        If Len(companyName) > 0 Then
            currentRecap = GetCompanyRecap(companyName, useOpenAI)
        Else
            currentRecap = "Invalid Applicant Name"
            LogEvt "ScoreLoop", lvlWARN, "Row " & i & ": Invalid/blank Applicant name."
        End If
        WriteResultsToArray dataArr, i, colIndices, scoreResult, currentRecap
        If i Mod 50 = 0 Or i = recordCount Then
            Application.StatusBar = "Calculating scores and fetching recaps (" & Format(i / recordCount, "0%") & " Complete)..."
            DoEvents
        End If
        If i Mod 100 = 0 Then
            LogEvt "ScoreLoop", lvlDETAIL, "Processed " & i & " of " & recordCount & " records (" & Format(i / recordCount, "0%") & ")"
        End If
    Next i
    LogEvt "ScoreLoop", lvlINFO, "Main processing loop complete."

    ' --- Write Processed Array Back to Sheet ---
    Application.StatusBar = "Writing results back to Excel sheet..."
    LogEvt "WriteBack", lvlINFO, "Writing " & recordCount & " rows back to table '" & tblData.Name & "'."
    tblData.DataBodyRange.Value = dataArr
    LogEvt "WriteBack", lvlINFO, "Array write complete."

    ' --- Apply Number Formats ---
    LogEvt "Formatting", lvlINFO, "Applying number formats."
    ApplyNumberFormats tblData

    ' --- Sort Table by DecisionDate ---
    Application.StatusBar = "Sorting data..."
    LogEvt "Sort", lvlINFO, "Sorting table by Decision Date (Descending)."
    SortDataTable tblData, "DecisionDate", xlDescending

    ' --- Save Updated Company Cache ---
    Application.StatusBar = "Saving company cache..."
    If Not wsCache Is Nothing And Not dictCache Is Nothing Then
        If dictCache.Count > 0 Then
            LogEvt "SaveCache", lvlINFO, "Saving " & dictCache.Count & " items to cache sheet '" & wsCache.Name & "'."
            Call SaveCompanyCache(wsCache)
        Else
            LogEvt "SaveCache", lvlINFO, "In-memory cache is empty, skipping save to sheet."
        End If
    End If

    ' --- Final Layout, Formatting & Visual Polish ---
    Application.StatusBar = "Applying final layout and formatting..."
    LogEvt "Formatting", lvlINFO, "Applying final layout and formatting."
    ' REVISED: Always call ReorganizeColumns first to ensure proper column order
    Call ReorganizeColumns(tblData) ' <<< USES CORRECTED VERSION
    Call FormatTableLook(wsData)    ' <<< USES CORRECTED VERSION
    Call FormatCategoryColors(tblData)
    Call CreateShortNamesAndComments(tblData)
    Call FreezeHeaderAndKeyCols(wsData)
    LogEvt "Formatting", lvlINFO, "Final formatting applied."

    ' --- Archive Month (if needed) ---
    Application.StatusBar = "Archiving month: " & targetMonthName & "..."
    If Not SheetExists(archiveSheetName) Then
        LogEvt "Archive", lvlINFO, "Starting archive creation for " & targetMonthName & "."
        Call ArchiveMonth(wsData, archiveSheetName)
    Else
        LogEvt "Archive", lvlWARN, "Archive sheet '" & archiveSheetName & "' already exists - skipping creation."
    End If

    ' --- Clean up duplicate connections created by sheet copy --- <<< ADDED
    Dim c As WorkbookConnection
    Dim baseConnectionName As String
    Dim originalConnection As WorkbookConnection ' To store ref if found

    ' Try to find the original connection (could have prefix or not)
    Set originalConnection = Nothing
    On Error Resume Next ' Handle if connection doesn't exist by either name
    Set originalConnection = ThisWorkbook.Connections("pgGet510kData") ' Try direct name first
    If originalConnection Is Nothing Then
        Set originalConnection = ThisWorkbook.Connections("Query - pgGet510kData") ' Try name with prefix
    End If
    On Error GoTo 0 ' Restore normal error handling

    If Not originalConnection Is Nothing Then
        baseConnectionName = originalConnection.Name ' Use the name of the connection found
        LogEvt "Cleanup", lvlINFO, "Checking for duplicate connections based on found connection: '" & baseConnectionName & "'"
        On Error Resume Next ' Ignore errors during loop/delete
        For Each c In ThisWorkbook.Connections
            ' Check if name starts with base name AND is not the original connection itself AND has " (" indicating a copy
            If c.Name <> baseConnectionName And c.Name Like baseConnectionName & " (*" Then
                LogEvt "Cleanup", lvlDETAIL, "Deleting duplicate connection: " & c.Name
                c.Delete
            End If
        Next c
        On Error GoTo 0 ' Restore error handling
    Else
         ' Could not find the original connection by either expected name pattern
         baseConnectionName = "pgGet510kData" ' Fallback base name for cleanup attempt
         LogEvt "Cleanup", lvlWARN, "Could not find original PQ connection by typical names. Attempting cleanup based on: '" & baseConnectionName & "'"
         On Error Resume Next ' Ignore errors during loop/delete
         For Each c In ThisWorkbook.Connections
             ' Check if name starts with base name and has " (" indicating a copy
             If c.Name Like baseConnectionName & " (*" Or c.Name Like "Query - " & baseConnectionName & " (*" Then
                 LogEvt "Cleanup", lvlDETAIL, "Deleting potential duplicate connection: " & c.Name
                 c.Delete
             End If
         Next c
         On Error GoTo 0 ' Restore error handling
    End If
    Set c = Nothing
    Set originalConnection = Nothing
    ' --- End of duplicate connection cleanup ---

    ' --- Completion Message ---
    Dim endTime As Double: endTime = Timer
    Dim elapsed As String: elapsed = Format(endTime - startTime, "0.00")
    LogEvt "ProcessEnd", lvlINFO, "Processing completed successfully.", "Records=" & recordCount & ", Elapsed=" & elapsed & "s"
    Application.StatusBar = "Processing complete for " & targetMonthName & "."
    MsgBox "Monthly 510(k) data processed and archived for " & targetMonthName & "." & vbCrLf & vbCrLf & _
           "Processed " & recordCount & " records in " & elapsed & " seconds.", vbInformation, "Processing Complete"

CleanExit:
    LogEvt "Cleanup", lvlINFO, "CleanExit reached. Releasing objects and restoring settings."
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    Set dictCache = Nothing: Set colIndices = Nothing
    Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing: Set wsLog = Nothing: Set tblData = Nothing
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    If Not Application.EnableEvents Then Application.EnableEvents = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Debug.Print Time & " - ProcessMonthly510k Finished. Objects released."
    FlushLogBuf
    Exit Sub

ProcessErrorHandler:
      Dim errNum As Long: errNum = Err.Number
      Dim errDesc As String: errDesc = Err.Description
      Dim errSource As String: errSource = Err.Source
      LogEvt "ProcessError", lvlERROR, "Error #" & errNum & " in " & errSource & ": " & errDesc
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      Debug.Print Time & " - FATAL ERROR #" & errNum & " in " & errSource & " (ProcessMonthly510k): " & errDesc
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      MsgBox "A critical error occurred during 510(k) processing:" & vbCrLf & vbCrLf & _
             "Error Number: " & errNum & vbCrLf & _
             "Module: mod_510k_Processor" & vbCrLf & _
             "Procedure: ProcessMonthly510k" & vbCrLf & _
             "Source: " & errSource & vbCrLf & _
             "Description: " & errDesc & vbCrLf & vbCrLf & _
             "Processing has been stopped.", vbCritical, "Processing Error"
      On Error Resume Next
      FlushLogBuf
      On Error GoTo 0
      Resume CleanExit
End Sub

'========================================================================
'  SortDataTable  –  generic ListObject sorter
'========================================================================
Public Sub SortDataTable(tbl As ListObject, _
                         colName As String, _
                         Optional sortOrder As XlSortOrder = xlAscending)
    On Error GoTo SortErr
    Dim loCol As ListColumn
    On Error Resume Next ' Handle column not found gracefully
    Set loCol = tbl.ListColumns(colName)
    On Error GoTo SortErr ' Restore error handler
    If loCol Is Nothing Then
        LogEvt "Sort", lvlWARN, "Column '" & colName & "' not found in table '" & tbl.Name & "' – sort skipped"
        Exit Sub
    End If
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=loCol.DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=sortOrder
        .Header = xlYes
        .MatchCase = False ' Default
        .Orientation = xlTopToBottom ' Default
        .SortMethod = xlPinYin ' Default
        .Apply
    End With
    LogEvt "Sort", lvlINFO, "Table '" & tbl.Name & "' sorted by " & _
           colName & IIf(sortOrder = xlDescending, " (Desc)", " (Asc)")
    Exit Sub
SortErr:
    LogEvt "Sort", lvlERROR, "Error sorting table '" & tbl.Name & "' by column '" & colName & "': " & Err.Description
End Sub


' ==========================================================================
' ===                CORE SCORING FUNCTION                         ===
' ==========================================================================
Private Function Calculate510kScore(dataArr As Variant, rowIdx As Long, ByVal cols As Object) As Variant
    ' Purpose: Calculates the 510(k) score based on various factors for a single record.
    ' Inputs:  dataArr - The 2D variant array holding all data.
    '          rowIdx - The current row number being processed in the array.
    '          cols - Dictionary mapping column names to their indices.
    ' Returns: A Variant array containing score components:
    '          Array(0=FinalScore, 1=Category, 2=AC_Wt, 3=PC_Wt, 4=KW_Wt, 5=ST_Wt,
    '                6=PT_Wt, 7=GL_Wt, 8=NF_Calc, 9=Synergy_Calc)

    ' --- Variable Declarations ---
    Dim AC As String, PC As String, DeviceName As String, Statement As String, SubmType As String, Country As String
    Dim ProcTimeDays As Variant, combinedText As String
    Dim AC_Wt As Double, PC_Wt As Double, KW_Wt As Double, ST_Wt As Double, PT_Wt As Double, GL_Wt As Double
    Dim NF_Calc As Double, Synergy_Calc As Double, Final_Score_Raw As Double
    Dim Category As String
    Dim HasHighValueKW As Boolean, IsCosmetic As Boolean, IsDiagnostic As Boolean, HasTherapeuticMention As Boolean
    Dim kw As Variant, kNum As String ' For loops and error logging

    ' --- Error Handling for this Function ---
    On Error GoTo ScoreErrorHandler

    ' --- 1. Extract Data Using Column Indices ---
    AC = SafeGetString(dataArr, rowIdx, cols, "AC")
    PC = SafeGetString(dataArr, rowIdx, cols, "PC")
    DeviceName = SafeGetString(dataArr, rowIdx, cols, "DeviceName")
    Statement = SafeGetString(dataArr, rowIdx, cols, "Statement")
    SubmType = SafeGetString(dataArr, rowIdx, cols, "SubmType")
    Country = UCase(SafeGetString(dataArr, rowIdx, cols, "Country"))
    ProcTimeDays = SafeGetVariant(dataArr, rowIdx, cols, "ProcTimeDays")
    combinedText = DeviceName & " " & Statement ' For keyword searching
    kNum = SafeGetString(dataArr, rowIdx, cols, "K_Number") ' For logging context

    ' --- 2. Calculate Individual Weights ---
    AC_Wt = GetWeightFromDict(dictACWeights, AC, DEFAULT_AC_WEIGHT)
    PC_Wt = GetWeightFromDict(dictPCWeights, PC, DEFAULT_PC_WEIGHT)
    ST_Wt = GetWeightFromDict(dictSTWeights, SubmType, DEFAULT_ST_WEIGHT)

    If IsNumeric(ProcTimeDays) Then
        Select Case CDbl(ProcTimeDays)
            Case Is > 172: PT_Wt = 0.65
            Case 162 To 172: PT_Wt = 0.6
            Case Else: PT_Wt = 0.5 ' Includes < 162 and non-positive values handled by IsNumeric
        End Select
    Else: PT_Wt = DEFAULT_PT_WEIGHT
    End If

    If Country = "US" Then GL_Wt = US_GL_WEIGHT Else GL_Wt = OTHER_GL_WEIGHT

    HasHighValueKW = CheckKeywords(combinedText, highValKeywordsList)
    If HasHighValueKW Then KW_Wt = HIGH_KW_WEIGHT Else KW_Wt = LOW_KW_WEIGHT

    ' --- 3. Negative Factors (NF) & Synergy Logic ---
    NF_Calc = 0: Synergy_Calc = 0
    IsCosmetic = CheckKeywords(combinedText, nfCosmeticKeywordsList)
    IsDiagnostic = CheckKeywords(combinedText, nfDiagnosticKeywordsList)
    HasTherapeuticMention = CheckKeywords(combinedText, therapeuticKeywordsList)

    If IsCosmetic And Not HasTherapeuticMention Then NF_Calc = NF_COSMETIC
    If IsDiagnostic And Not HasTherapeuticMention Then
        If NF_Calc = 0 Then NF_Calc = NF_DIAGNOSTIC Else NF_Calc = NF_Calc + NF_DIAGNOSTIC ' Confirm additive logic if needed
    End If

    If (AC = "OR" Or AC = "NE") And HasHighValueKW Then Synergy_Calc = SYNERGY_BONUS

    ' --- 4. Final Score Calculation ---
    Final_Score_Raw = (AC_Wt + PC_Wt + KW_Wt + ST_Wt + PT_Wt + GL_Wt + NF_Calc + Synergy_Calc) / 6 ' Confirm divisor logic
    If Final_Score_Raw < 0 Then Final_Score_Raw = 0

    ' --- 5. Determine Category ---
    Select Case Final_Score_Raw
        Case Is > 0.6: Category = "High"
        Case 0.5 To 0.6: Category = "Moderate"
        Case 0.4 To 0.499999999999: Category = "Low"
        Case Else: Category = "Almost None"
    End Select

    ' --- 6. Return Results ---
    Calculate510kScore = Array(Final_Score_Raw, Category, AC_Wt, PC_Wt, KW_Wt, ST_Wt, PT_Wt, GL_Wt, NF_Calc, Synergy_Calc)
    Exit Function

ScoreErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    LogEvt "ScoreError", lvlERROR, "Error scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc, "AC=" & AC & ", PC=" & PC ' Use LogEvt
    Debug.Print Time & " - ERROR scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc
    Calculate510kScore = Array(0, "Error", 0, 0, 0, 0, 0, 0, 0, 0) ' Return error state array
End Function


' ==========================================================================
' ===                COMPANY RECAP & CACHING FUNCTIONS                   ===
' ==========================================================================
Private Function GetCompanyRecap(companyName As String, useOpenAI As Boolean) As String
    Dim finalRecap As String
    If dictCache Is Nothing Then Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare
    If Len(Trim(companyName)) = 0 Then GetCompanyRecap = "Invalid Applicant Name": Exit Function

    If dictCache.Exists(companyName) Then
        finalRecap = dictCache(companyName)
        LogEvt "CacheCheck", lvlDETAIL, "Memory Cache HIT.", "Company=" & companyName ' Changed level to DETAIL
    Else
        LogEvt "CacheCheck", lvlDETAIL, "Memory Cache MISS.", "Company=" & companyName ' Changed level to DETAIL
        finalRecap = DEFAULT_RECAP_TEXT
        If useOpenAI Then
            Dim openAIResult As String
            LogEvt "OpenAI", lvlINFO, "Attempting OpenAI call.", "Company=" & companyName
            openAIResult = GetCompanyRecapOpenAI(companyName)
            If openAIResult <> "" And Not LCase(openAIResult) Like "error:*" Then
                finalRecap = openAIResult
                LogEvt "OpenAI", lvlINFO, "OpenAI SUCCESS.", "Company=" & companyName
            Else
                LogEvt "OpenAI", IIf(LCase(openAIResult) Like "error:*", lvlERROR, lvlWARN), "OpenAI Failed or Skipped. Result: " & openAIResult, "Company=" & companyName
            End If
        Else
             LogEvt "OpenAI", lvlINFO, "OpenAI call skipped (Not Maintainer).", "Company=" & companyName
        End If
        On Error Resume Next
        dictCache(companyName) = finalRecap
        If Err.Number <> 0 Then LogEvt "CacheUpdate", lvlERROR, "Error adding '" & companyName & "' to memory cache: " & Err.Description: Err.Clear
        On Error GoTo 0
    End If
    GetCompanyRecap = finalRecap
End Function

Private Function GetCompanyRecapOpenAI(companyName As String) As String
    Dim apiKey As String, result As String, http As Object, url As String, jsonPayload As String, jsonResponse As String
    GetCompanyRecapOpenAI = "" ' Default return

    If Not IsMaintainerUser() Then
         LogEvt "OpenAI_Skip", lvlINFO, "Skipped OpenAI Call: Not Maintainer User.", "Company=" & companyName
        Exit Function
    End If

    apiKey = GetAPIKey()
    If apiKey = "" Then
         LogEvt "OpenAI_Skip", lvlERROR, "Skipped OpenAI Call: API Key Not Found/Configured.", "Company=" & companyName
        GetCompanyRecapOpenAI = "Error: API Key Not Configured"
        Exit Function
    End If

    On Error GoTo OpenAIErrorHandler

    url = OPENAI_API_URL
    Dim modelName As String: modelName = OPENAI_MODEL
    Dim systemPrompt As String, userPrompt As String
    systemPrompt = "You are an analyst summarizing medical device related companies based *only* on publicly available information. " & _
                   "Provide a *neutral*, *very concise* (1 sentence ideally, 2 max) summary of the company '" & Replace(companyName, """", "'") & "' " & _
                   "identifying its primary business sector or main product type (e.g., orthopedics, diagnostics, surgical tools, contract manufacturer). " & _
                   "If unsure or company is generic, state 'General medical device company'. Avoid speculation or marketing language."
    userPrompt = "Summarize: " & companyName

    jsonPayload = "{""model"": """ & modelName & """, ""messages"": [" & _
                  "{""role"": ""system"", ""content"": """ & JsonEscape(systemPrompt) & """}," & _
                  "{""role"": ""user"", ""content"": """ & JsonEscape(userPrompt) & """}" & _
                  "], ""temperature"": 0.3, ""max_tokens"": " & OPENAI_MAX_TOKENS & "}"

    LogEvt "OpenAI_HTTP", lvlDETAIL, "Sending request...", "Company=" & companyName & ", Model=" & modelName ' Changed level
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setTimeouts OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send jsonPayload

    LogEvt "OpenAI_HTTP", lvlDETAIL, "Response Received.", "Company=" & companyName & ", Status=" & http.Status ' Changed level

    If http.Status = 200 Then
        jsonResponse = http.responseText
        Const CONTENT_TAG As String = """content"":"""
        Dim contentStart As Long, contentEnd As Long, searchStart As Long
        searchStart = InStr(1, jsonResponse, """role"":""assistant""")
        If searchStart > 0 Then
            contentStart = InStr(searchStart, jsonResponse, CONTENT_TAG)
            If contentStart > 0 Then
                contentStart = contentStart + Len(CONTENT_TAG)
                contentEnd = InStr(contentStart, jsonResponse, """")
                If contentEnd > contentStart Then
                    result = Mid$(jsonResponse, contentStart, contentEnd - contentStart)
                    result = JsonUnescape(result)
                Else: result = "Error: Parse Fail (End Quote)": LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
                End If
            Else
                If InStr(1, jsonResponse, """error""", vbTextCompare) > 0 Then
                     result = "Error: API returned error object."
                     LogEvt "OpenAI_APIError", lvlERROR, result, Left(jsonResponse, 500)
                Else
                     result = "Error: Parse Fail (Start Tag)"
                     LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
                End If
            End If
        Else: result = "Error: Parse Fail (No Assistant Role)": LogEvt "OpenAI_Parse", lvlERROR, result, Left(jsonResponse, 500)
        End If
    Else
        result = "Error: API Call Failed - Status " & http.Status & " - " & http.statusText
         LogEvt "OpenAI_HTTP", lvlERROR, result, Left(http.responseText, 500)
    End If

    Set http = Nothing
    If Len(result) > RECAP_MAX_LEN Then result = Left$(result, RECAP_MAX_LEN) & "..."
    GetCompanyRecapOpenAI = Trim(result)
    Exit Function

OpenAIErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     LogEvt "OpenAI_VBAError", lvlERROR, "VBA Exception during OpenAI Call: " & errDesc, "Company=" & companyName
    GetCompanyRecapOpenAI = "Error: VBA Exception - " & errDesc
    If Not http Is Nothing Then Set http = Nothing
End Function

Private Sub LoadCompanyCache(wsCache As Worksheet)
    Dim lastRow As Long, i As Long, cacheData As Variant, loadedCount As Long
    Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare
    On Error Resume Next
    lastRow = wsCache.Cells(wsCache.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Or lastRow < 2 Then GoTo ExitLoadCache
    On Error GoTo CacheLoadError
    cacheData = wsCache.Range("A2:B" & lastRow).Value2 ' Read CompanyName and RecapText only
    If Err.Number <> 0 Then GoTo CacheLoadError

    If IsArray(cacheData) Then
        For i = 1 To UBound(cacheData, 1)
            Dim k As String: k = Trim(CStr(cacheData(i, 1)))
            Dim v As String: v = CStr(cacheData(i, 2))
            If Len(k) > 0 Then
                If Not dictCache.Exists(k) Then dictCache.Add k, v Else dictCache(k) = v
            End If
        Next i
    ElseIf lastRow = 2 Then ' Handle single data row case explicitly
         Dim kS As String: kS = Trim(CStr(wsCache.Range("A2").Value2))
         Dim vS As String: vS = CStr(wsCache.Range("B2").Value2)
         If Len(kS) > 0 Then dictCache(kS) = vS
    End If
ExitLoadCache:
    loadedCount = dictCache.Count
     LogEvt "LoadCache", IIf(Err.Number <> 0 And loadedCount = 0, lvlWARN, lvlINFO), "Loaded " & loadedCount & " items into memory cache.", IIf(Err.Number <> 0 And loadedCount = 0, "Sheet might be empty or error occurred finding data.", "")
    Err.Clear
    Exit Sub
CacheLoadError:
     LogEvt "LoadCache", lvlERROR, "Error reading cache data from sheet: " & Err.Description
     Err.Clear
     Resume ExitLoadCache
End Sub

Private Sub SaveCompanyCache(wsCache As Worksheet)
    Dim key As Variant, i As Long, outputArr() As Variant, saveCount As Long
    If dictCache Is Nothing Or dictCache.Count = 0 Then LogEvt "SaveCache", lvlINFO, "In-memory cache empty, skipping save.": Exit Sub

    On Error GoTo CacheSaveError
    saveCount = dictCache.Count
    ReDim outputArr(1 To saveCount, 1 To 3) ' CompanyName, RecapText, LastUpdated
    i = 1
    For Each key In dictCache.Keys
        outputArr(i, 1) = key
        outputArr(i, 2) = dictCache(key)
        outputArr(i, 3) = Now
        i = i + 1
    Next key

    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With wsCache
        .Range("A2:C" & .Rows.Count).ClearContents
        If saveCount > 0 Then
            .Range("A2").Resize(saveCount, 3).Value = outputArr
            .Range("C2").Resize(saveCount, 1).NumberFormat = "m/d/yyyy h:mm AM/PM"
            .Columns("A:C").AutoFit ' Autofit columns after writing
        End If
    End With
     LogEvt "SaveCache", lvlINFO, "Saved " & saveCount & " items to cache sheet."
    GoTo CacheSaveExit

CacheSaveError:
     LogEvt "SaveCache", lvlERROR, "Error saving cache to sheet '" & wsCache.Name & "': " & Err.Description
    MsgBox "Error saving company cache to sheet '" & wsCache.Name & "': " & Err.Description, vbExclamation, "Cache Save Error"
CacheSaveExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ==========================================================================
' ===                WEIGHTS & KEYWORDS LOADING FUNCTIONS                ===
' ==========================================================================
Private Function LoadWeightsAndKeywords(wsWeights As Worksheet) As Boolean
    Dim success As Boolean: success = True
    On Error GoTo LoadErrorHandler
    LogEvt "LoadParams", lvlINFO, "Attempting to load weights and keywords from sheet: " & wsWeights.Name

    Set dictACWeights = LoadTableToDict(wsWeights, "tblACWeights")
    Set dictSTWeights = LoadTableToDict(wsWeights, "tblSTWeights")
    Set dictPCWeights = LoadTableToDict(wsWeights, "tblPCWeights")
    Set highValKeywordsList = LoadTableToList(wsWeights, "tblKeywords")
    Set nfCosmeticKeywordsList = LoadTableToList(wsWeights, "tblNFCosmeticKeywords")
    Set nfDiagnosticKeywordsList = LoadTableToList(wsWeights, "tblNFDiagnosticKeywords")
    Set therapeuticKeywordsList = LoadTableToList(wsWeights, "tblTherapeuticKeywords")

    If dictACWeights Is Nothing Or dictSTWeights Is Nothing Or highValKeywordsList Is Nothing Then
         LogEvt "LoadParams", lvlERROR, "Critical failure: Could not load AC/ST weights or HighValue Keywords." ' Changed level
        GoTo LoadErrorCritical
    End If

    LogEvt "LoadParams", IIf(dictACWeights.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & dictACWeights.Count & " AC Weights."
    LogEvt "LoadParams", IIf(dictSTWeights.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & dictSTWeights.Count & " ST Weights."
    LogEvt "LoadParams", IIf(dictPCWeights Is Nothing Or dictPCWeights.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & IIf(dictPCWeights Is Nothing, 0, dictPCWeights.Count) & " PC Weights (Optional)."
    LogEvt "LoadParams", IIf(highValKeywordsList.Count = 0, lvlWARN, lvlDETAIL), "Loaded " & highValKeywordsList.Count & " HighVal Keywords."
    LogEvt "LoadParams", IIf(nfCosmeticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & nfCosmeticKeywordsList.Count & " Cosmetic Keywords (Optional)."
    LogEvt "LoadParams", IIf(nfDiagnosticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & nfDiagnosticKeywordsList.Count & " Diagnostic Keywords (Optional)."
    LogEvt "LoadParams", IIf(therapeuticKeywordsList.Count = 0, lvlINFO, lvlDETAIL), "Loaded " & therapeuticKeywordsList.Count & " Therapeutic Keywords (Optional)."

    LoadWeightsAndKeywords = True
    Exit Function

LoadErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     LogEvt "LoadParams", lvlWARN, "Non-critical error loading one or more weight/keyword tables: " & errDesc & ". Defaults will be used.", "Sheet=" & wsWeights.Name
    MsgBox "Warning: Error loading weights/keywords from '" & wsWeights.Name & "'. " & vbCrLf & _
           "Check tables exist and are named correctly (e.g., tblACWeights, tblKeywords, tblNFCosmeticKeywords etc.)." & vbCrLf & _
           "Default values will be used where possible.", vbExclamation, "Load Warning"
    If dictACWeights Is Nothing Then Set dictACWeights = CreateObject("Scripting.Dictionary")
    If dictSTWeights Is Nothing Then Set dictSTWeights = CreateObject("Scripting.Dictionary")
    If dictPCWeights Is Nothing Then Set dictPCWeights = CreateObject("Scripting.Dictionary")
    If highValKeywordsList Is Nothing Then Set highValKeywordsList = New Collection
    If nfCosmeticKeywordsList Is Nothing Then Set nfCosmeticKeywordsList = New Collection
    If nfDiagnosticKeywordsList Is Nothing Then Set nfDiagnosticKeywordsList = New Collection
    If therapeuticKeywordsList Is Nothing Then Set therapeuticKeywordsList = New Collection
    LoadWeightsAndKeywords = True
    Exit Function

LoadErrorCritical:
    MsgBox "Critical Error: Could not create necessary objects for AC/ST weights or HighValue Keywords. Processing cannot continue.", vbCritical, "Load Failure"
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    LoadWeightsAndKeywords = False
End Function

Private Function LoadTableToDict(ws As Worksheet, tableName As String) As Object ' Scripting.Dictionary
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, key As String, val As Variant

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    If Err.Number <> 0 Then GoTo ExitLoadTableDict
    If tbl.ListRows.Count = 0 Then GoTo ExitLoadTableDict
    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Or dataRange.Columns.Count < 2 Then GoTo ExitLoadTableDict
    dataArr = dataRange.Value2
    If Err.Number <> 0 Then GoTo ExitLoadTableDict
    On Error GoTo 0 ' Restore error handling after checks

    If IsArray(dataArr) Then
        For i = 1 To UBound(dataArr, 1)
            key = Trim(CStr(dataArr(i, 1)))
            val = dataArr(i, 2)
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, val Else dict(key) = val
            End If
        Next i
    ElseIf Not IsEmpty(dataArr) And tbl.ListRows.Count = 1 And tbl.ListColumns.Count >= 2 Then ' Handle single row table explicitly
         key = Trim(CStr(tbl.DataBodyRange.Cells(1, 1).Value2))
         val = tbl.DataBodyRange.Cells(1, 2).Value2
         If Len(key) > 0 Then If Not dict.Exists(key) Then dict.Add key, val Else dict(key) = val
    End If

ExitLoadTableDict:
    If Err.Number <> 0 Then
         LogEvt "LoadHelper", lvlWARN, "Error loading table '" & tableName & "' to Dict: " & Err.Description
        Debug.Print Time & " - Error loading table '" & tableName & "' to Dict: " & Err.Description: Err.Clear
    End If
    Set LoadTableToDict = dict
End Function

Private Function LoadTableToList(ws As Worksheet, tableName As String) As Collection
    Dim coll As New Collection, tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, item As String
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    If Err.Number <> 0 Then GoTo ExitLoadTableList
    If tbl.ListRows.Count = 0 Then GoTo ExitLoadTableList
    Set dataRange = tbl.ListColumns(1).DataBodyRange
    If dataRange Is Nothing Then GoTo ExitLoadTableList
    dataArr = dataRange.Value2
    If Err.Number <> 0 Then GoTo ExitLoadTableList
    On Error GoTo 0 ' Restore error handling

    If IsArray(dataArr) Then
        For i = 1 To UBound(dataArr, 1)
            item = Trim(CStr(dataArr(i, 1)))
            If Len(item) > 0 Then On Error Resume Next: coll.Add item, item: On Error GoTo 0 ' Add unique items by key
        Next i
    ElseIf Not IsEmpty(dataArr) Then ' Handle single row data
        item = Trim(CStr(dataArr))
        If Len(item) > 0 Then On Error Resume Next: coll.Add item, item: On Error GoTo 0
    End If

ExitLoadTableList:
    If Err.Number <> 0 Then
         LogEvt "LoadHelper", lvlWARN, "Error loading table '" & tableName & "' to List: " & Err.Description
        Debug.Print Time & " - Error loading table '" & tableName & "' to List: " & Err.Description: Err.Clear
    End If
    Set LoadTableToList = coll
End Function

Private Function GetWeightFromDict(dict As Object, key As String, defaultWeight As Double) As Double
    If dict Is Nothing Then GetWeightFromDict = defaultWeight: Exit Function
    On Error Resume Next ' Handle potential type mismatch if value isn't numeric
    If dict.Exists(key) Then
        GetWeightFromDict = CDbl(dict(key)) ' Explicitly convert to Double
        If Err.Number <> 0 Then GetWeightFromDict = defaultWeight: Err.Clear ' Use default if conversion fails
    Else
        GetWeightFromDict = defaultWeight
    End If
    On Error GoTo 0
End Function

Private Function CheckKeywords(textToCheck As String, keywordColl As Collection) As Boolean
    Dim kw As Variant
    CheckKeywords = False
    If keywordColl Is Nothing Or keywordColl.Count = 0 Or Len(Trim(textToCheck)) = 0 Then Exit Function
    On Error Resume Next
    For Each kw In keywordColl
        If InStr(1, textToCheck, CStr(kw), vbTextCompare) > 0 Then CheckKeywords = True: Exit For
    Next kw
    On Error GoTo 0
End Function

' ==========================================================================
' ===                   COLUMN MANAGEMENT FUNCTIONS                      ===
' ==========================================================================
Private Function AddScoreColumnsIfNeeded(tblData As ListObject) As Boolean
    Dim requiredCols As Variant, colName As Variant, col As ListColumn, addedCol As Boolean: addedCol = False
    On Error GoTo AddColErrorHandler
    requiredCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")
    LogEvt "Columns", lvlDETAIL, "Verifying existence of calculated columns..."
    Dim currentHeaders As String, hCell As Range
    For Each hCell In tblData.HeaderRowRange: currentHeaders = currentHeaders & "|" & UCase(Trim(hCell.Value)) & "|": Next hCell

    For Each colName In requiredCols
        If InStr(1, currentHeaders, "|" & UCase(colName) & "|", vbTextCompare) = 0 Then
            On Error Resume Next
            Set col = tblData.ListColumns.Add()
            If Err.Number <> 0 Then
                LogEvt "Columns", lvlERROR, "Failed to add column '" & colName & "'. Error: " & Err.Description
                GoTo AddColErrorHandler ' Treat failure to add column as critical
            End If
            col.Name = colName
            addedCol = True
            LogEvt "Columns", lvlINFO, "Added missing column: " & colName
            On Error GoTo AddColErrorHandler ' Restore handler for next iteration
        End If
        Set col = Nothing
    Next colName

    AddScoreColumnsIfNeeded = True
    Exit Function
AddColErrorHandler:
     LogEvt "Columns", lvlERROR, "Error checking/adding columns to table '" & tblData.Name & "': " & Err.Description
    MsgBox "Error verifying or adding required columns to table '" & tblData.Name & "': " & vbCrLf & Err.Description, vbCritical, "Column Setup Error"
    AddScoreColumnsIfNeeded = False
End Function

Private Sub WriteResultsToArray(ByRef dataArr As Variant, ByVal rowIdx As Long, ByVal cols As Object, ByVal scoreResult As Variant, ByVal recap As String)
    On Error Resume Next ' Suppress errors during write
    dataArr(rowIdx, cols("Final_Score")) = scoreResult(0)
    dataArr(rowIdx, cols("Category")) = scoreResult(1)
    dataArr(rowIdx, cols("AC_Wt")) = scoreResult(2)
    dataArr(rowIdx, cols("PC_Wt")) = scoreResult(3)
    dataArr(rowIdx, cols("KW_Wt")) = scoreResult(4)
    dataArr(rowIdx, cols("ST_Wt")) = scoreResult(5)
    dataArr(rowIdx, cols("PT_Wt")) = scoreResult(6)
    dataArr(rowIdx, cols("GL_Wt")) = scoreResult(7)
    dataArr(rowIdx, cols("NF_Calc")) = scoreResult(8)
    dataArr(rowIdx, cols("Synergy_Calc")) = scoreResult(9)
    dataArr(rowIdx, cols("Score_Percent")) = scoreResult(0) ' Store raw decimal
    dataArr(rowIdx, cols("CompanyRecap")) = recap
    If Err.Number <> 0 Then
         LogEvt "WriteToArray", lvlERROR, "Error writing results to array row " & rowIdx & ": " & Err.Description
        Debug.Print Time & " - Error writing results to array row " & rowIdx & ": " & Err.Description: Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function GetColumnIndices(headerRange As Range) As Object ' Scripting.Dictionary
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim cell As Range, colNum As Long, missingCols As String, h As String
    colNum = 1
    Dim requiredCols As Variant
    requiredCols = Array("K_Number", "Applicant", "Contact", "DeviceName", "DecisionDate", "DateReceived", "ProcTimeDays", "AC", "PC", "SubmType", "City", "State", "Country", "Statement", "FDA_Link", _
                         "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")

    For Each cell In headerRange.Cells
        h = Trim(cell.Value)
        If Len(h) > 0 Then
            If Not dict.Exists(h) Then
                dict.Add h, colNum
            Else
                ' Log duplicate header only once if it's immediately adjacent
                ' If colNum = dict(h) + 1 Then ' This logic might be flawed if columns are reordered before mapping
                     LogEvt "ColumnMap", lvlWARN, "Duplicate header found and may cause issues: '" & h & "' at column " & colNum & " (and earlier)"
                    Debug.Print Time & " - Warning: Duplicate header found: '" & h & "' at column " & colNum
                ' End If
            End If
        End If
        colNum = colNum + 1
    Next cell

    Dim reqCol As Variant
    For Each reqCol In requiredCols
        If Not dict.Exists(reqCol) Then
            missingCols = missingCols & vbCrLf & " - " & reqCol
        End If
    Next reqCol

    If Len(missingCols) > 0 Then
         LogEvt "ColumnMap", lvlERROR, "Required columns missing in table header:" & Replace(missingCols, vbCrLf, ", ")
        MsgBox "Error: The following required columns were not found in sheet '" & headerRange.Parent.Name & "':" & missingCols & vbCrLf & "Please ensure Power Query output and VBA column additions match.", vbCritical, "Missing Columns"
        Set GetColumnIndices = Nothing
    Else
         LogEvt "ColumnMap", lvlINFO, "Column indices mapped successfully for " & dict.Count & " columns."
        Set GetColumnIndices = dict
    End If
End Function

' ==========================================================================
' ===             LAYOUT, FORMATTING & PRESENTATION FUNCTIONS            ===
' ==========================================================================

' <<< USING CUT/INSERT VERSION FOR COMPATIBILITY >>>
Private Sub ReorganizeColumns(tbl As ListObject)
    ' Purpose: Rearranges table columns into the final desired presentation order using Cut/Insert for broad compatibility.
    ' Note: Run AFTER all data processing and BEFORE formatting that relies on column position. Loop runs Right-to-Left.

    On Error GoTo ReorgErrorHandler

    ' *** Define the desired final column order ***
    Dim targetOrder As Variant: targetOrder = Array( _
        "K_Number", "DecisionDate", "ProcTimeDays", "Applicant", "DeviceName", "Contact", _
        "CompanyRecap", "Score_Percent", "Category", _
        "FDA_Link", "DateReceived", "AC", "PC", "SubmType", "City", "State", "Country", "Statement", _
        "Final_Score", "NF_Calc", "Synergy_Calc", "AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt" _
        ) ' <<< Columns listed first as requested, others follow >>>

    Dim targetPosition As Long
    Dim currentPosition As Long
    Dim col As ListColumn
    Dim currentColName As String
    Dim i As Long ' Loop counter

    Application.ScreenUpdating = False ' Turn off screen updating for performance
    LogEvt "Formatting", lvlDETAIL, "Starting column reorganization (using Cut/Insert method, R-to-L)..."

    ' <<< CORRECTED LOOP: Runs Right-to-Left (UBound to LBound) >>>
    For i = UBound(targetOrder) To LBound(targetOrder) Step -1
        currentColName = targetOrder(i)
        targetPosition = i + 1                ' 1-based index for columns

        Set col = Nothing ' Reset column object for safety
        On Error Resume Next                  ' Silently ignore if column name not found in table
        Set col = tbl.ListColumns(currentColName)
        On Error GoTo ReorgErrorHandler       ' Restore normal error handling for other potential errors

        If Not col Is Nothing Then
            ' Column found in the table
            currentPosition = col.Index ' Get its current position
            ' Only move if it's not already in the target position
            If currentPosition <> targetPosition Then
                ' Use Cut and Insert for universal compatibility
                Application.EnableEvents = False ' Prevent events during manipulation
                col.Range.Cut
                ' Insert before the column currently at the target position
                tbl.HeaderRowRange.Cells(1, targetPosition).EntireColumn.Insert Shift:=xlToRight
                Application.EnableEvents = True ' Re-enable events
                LogEvt "Formatting", lvlDETAIL, "Moved (Cut/Insert) column '" & currentColName & "' from " & currentPosition & " to " & targetPosition
            End If
        Else
            ' Column defined in targetOrder was not found in the table
            LogEvt "Formatting", lvlWARN, "Column '" & currentColName & "' not found during reorganization. Skipping."
        End If
        Set col = Nothing ' Release object for next iteration
    Next i

    Application.CutCopyMode = False ' Clear clipboard indicator after loop
    Application.ScreenUpdating = True ' Restore screen updating
    LogEvt "Formatting", lvlINFO, "Column reorganization complete."
    Exit Sub ' Normal exit

ReorgErrorHandler:
    Application.ScreenUpdating = True ' Ensure screen updating is restored on error
    Application.EnableEvents = True   ' Ensure events are re-enabled on error
    Application.CutCopyMode = False   ' Clear clipboard indicator on error
    LogEvt "Formatting", lvlERROR, "Error during column reorganization: " & Err.Description
    MsgBox "Error occurred during column reorganization: " & Err.Description, vbCritical, "Column Reorder Error"
    ' Exits implicitly after error

End Sub


' <<< USING VERSION WITH BLUE HEADER / BLACK BORDERS >>>
Private Sub FormatTableLook(ws As Worksheet)
    ' Purpose: Applies consistent table style, alignment, widths, borders, and specific header formatting.
    On Error GoTo FormatLookErrorHandler
    Dim tbl As ListObject
    Dim listCol As ListColumn
    Dim centerCols As Variant
    Dim wideCols As Variant
    Dim colName As Variant

    ' --- CORRECTED START ---
    ' Ensure there's a table on the sheet
    If ws.ListObjects.Count = 0 Then
        LogEvt "Formatting", lvlWARN, "No table found on sheet '" & ws.Name & "' for FormatTableLook."
        Exit Sub ' Exit if no table found
    End If
    Set tbl = ws.ListObjects(1) ' Assumes the first table is the target
    ' --- END CORRECTED START ---

    ' Define columns for specific formatting
    centerCols = Array("ProcTimeDays", "AC", "PC", "Category") ' Columns to center align
    wideCols = Array("DeviceName", "Applicant", "CompanyRecap") ' Columns needing more width

    LogEvt "Formatting", lvlDETAIL, "Applying table style, alignment, widths, borders for table '" & tbl.Name & "'..." ' Added table name to log

    ' Apply Base Table Style (optional, header formatting below will override header style)
    ' Setting to "" ensures no style conflicts with manual formatting
    tbl.TableStyle = "" ' Use empty string "" to remove base table style

    ' *** ADDED: Explicit Header Formatting ***
    With tbl.HeaderRowRange
        .Interior.Color = RGB(31, 78, 121)   ' Deep Excel blue (adjust RGB if needed)
        .Font.Color = vbWhite                ' White text
        .Font.Bold = True                    ' Bold text
        .HorizontalAlignment = xlCenter      ' Center header text
        .VerticalAlignment = xlCenter        ' Center vertically
    End With

    ' Center specific data columns (adjust array as needed)
    For Each colName In centerCols
        On Error Resume Next ' Ignore if column doesn't exist
        Set listCol = Nothing ' Reset object
        Set listCol = tbl.ListColumns(colName)
        If Not listCol Is Nothing Then
            listCol.DataBodyRange.HorizontalAlignment = xlCenter
        End If
        On Error GoTo FormatLookErrorHandler ' Restore handler
    Next colName
    Set listCol = Nothing ' Release object

    ' Autofit all columns first to get a baseline
    On Error Resume Next
    tbl.Range.Columns.AutoFit
    On Error GoTo FormatLookErrorHandler

    ' Set specific widths for wider columns after autofit (adjust widths as needed)
    On Error Resume Next
    If Not tbl.ListColumns("DeviceName") Is Nothing Then tbl.ListColumns("DeviceName").Range.ColumnWidth = 45
    If Not tbl.ListColumns("Applicant") Is Nothing Then tbl.ListColumns("Applicant").Range.ColumnWidth = 30
    If Not tbl.ListColumns("CompanyRecap") Is Nothing Then tbl.ListColumns("CompanyRecap").Range.ColumnWidth = 50
    On Error GoTo FormatLookErrorHandler

    ' *** UPDATED: Apply thin BLACK borders to ALL cells in the table range ***
    With tbl.Range.Borders
        .LineStyle = xlContinuous ' Ensure all borders are continuous lines
        .Weight = xlThin          ' Set line weight to thin
        .Color = vbBlack          ' Set border color to black
    End With
    ' Ensure inside borders are also set
    On Error Resume Next ' Apply to inside borders as well
    With tbl.Range.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = vbBlack
    End With
    With tbl.Range.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = vbBlack
    End With
    On Error GoTo FormatLookErrorHandler ' Restore handler

    LogEvt "Formatting", lvlDETAIL, "Applied FormatTableLook with custom header and black borders."
    Exit Sub

FormatLookErrorHandler:
    LogEvt "Formatting", lvlERROR, "Error applying table look formatting on sheet '" & ws.Name & "': " & Err.Description
    Debug.Print Time & " - Error applying table look formatting on sheet '" & ws.Name & "': " & Err.Description
    MsgBox "Error applying table formatting: " & Err.Description, vbExclamation

End Sub


Private Sub FormatCategoryColors(tblData As ListObject)
    Dim rng As Range
    On Error Resume Next
    Set rng = tblData.ListColumns("Category").DataBodyRange
    If rng Is Nothing Then Exit Sub
    On Error GoTo CatColorError
    LogEvt "Formatting", lvlDETAIL, "Applying category conditional formatting..."
    rng.FormatConditions.Delete
    Dim highColor As Long: highColor = RGB(198, 239, 206)
    Dim modColor As Long: modColor = RGB(255, 235, 156)
    Dim lowColor As Long: lowColor = RGB(255, 221, 179)
    Dim noneColor As Long: noneColor = RGB(242, 242, 242)
    Dim errorColor As Long: errorColor = RGB(255, 199, 206)
    ApplyCondColor rng, "High", highColor
    ApplyCondColor rng, "Moderate", modColor
    ApplyCondColor rng, "Low", lowColor
    ApplyCondColor rng, "Almost None", noneColor
    ApplyCondColor rng, "Error", errorColor
    LogEvt "Formatting", lvlDETAIL, "Applied category colors."
    Exit Sub
CatColorError:
    LogEvt "Formatting", lvlERROR, "Error applying category colors: " & Err.Description
    Debug.Print Time & " - Error applying category colors: " & Err.Description
End Sub

Private Sub ApplyCondColor(rng As Range, categoryText As String, fillColor As Long)
    With rng.FormatConditions.Add(Type:=xlTextString, String:=categoryText, TextOperator:=xlEqual)
        .Interior.Color = fillColor
        .Font.Color = IIf(GetBrightness(fillColor) < 130, vbWhite, vbBlack)
        .StopIfTrue = False
    End With
End Sub

Private Function GetBrightness(clr As Long) As Double
    On Error Resume Next
    GetBrightness = ((clr Mod 256) * 0.299) + (((clr \ 256) Mod 256) * 0.587) + (((clr \ 65536) Mod 256) * 0.114)
    If Err.Number <> 0 Then GetBrightness = 128
    On Error GoTo 0
End Function

Private Sub ApplyNumberFormats(tblData As ListObject)
    On Error Resume Next ' Ignore errors if a column doesn't exist
    LogEvt "Formatting", lvlDETAIL, "Applying number formats..."
    If Not tblData.ListColumns("Score_Percent") Is Nothing Then tblData.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
    If Not tblData.ListColumns("Final_Score") Is Nothing Then tblData.ListColumns("Final_Score").DataBodyRange.NumberFormat = "0.000"
    Dim scoreWtCols As Variant: scoreWtCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc")
    Dim colName As Variant
    For Each colName In scoreWtCols
        If Not tblData.ListColumns(colName) Is Nothing Then tblData.ListColumns(colName).DataBodyRange.NumberFormat = "0.00"
    Next colName
    If Not tblData.ListColumns("ProcTimeDays") Is Nothing Then tblData.ListColumns("ProcTimeDays").DataBodyRange.NumberFormat = "0"
    LogEvt "Formatting", lvlDETAIL, "Number formats applied."
    On Error GoTo 0
End Sub

Private Sub CreateShortNamesAndComments(tblData As ListObject)
    Dim devNameCol As ListColumn, devNameRange As Range, cell As Range
    Dim originalName As String, shortName As String
    On Error Resume Next
    Set devNameCol = tblData.ListColumns("DeviceName")
    If devNameCol Is Nothing Then Exit Sub
    Set devNameRange = devNameCol.DataBodyRange
    If devNameRange Is Nothing Then Exit Sub
    On Error GoTo ShortNameErrorHandler
    LogEvt "Formatting", lvlDETAIL, "Applying smart short names and comments..."
    Application.ScreenUpdating = False
    For Each cell In devNameRange.Cells
        If Not cell.Comment Is Nothing Then cell.Comment.Delete
        originalName = Trim(CStr(cell.Value))
        If Len(originalName) > SHORT_NAME_MAX_LEN Then
            If InStr(1, originalName, "(") > 10 Then
                shortName = Trim(Left(originalName, InStr(1, originalName, "(") - 1))
            ElseIf InStr(1, originalName, ";") > 10 Then
                shortName = Trim(Left(originalName, InStr(1, originalName, ";") - 1)) & "..."
            ElseIf InStr(1, originalName, ",") > 25 Then
                shortName = Trim(Left(originalName, InStr(1, originalName, ",") - 1)) & "..."
            Else
                shortName = Left$(originalName, SHORT_NAME_MAX_LEN - Len(SHORT_NAME_ELLIPSIS)) & SHORT_NAME_ELLIPSIS
            End If
            If Len(shortName) < 10 Then
                shortName = Left$(originalName, SHORT_NAME_MAX_LEN - Len(SHORT_NAME_ELLIPSIS)) & SHORT_NAME_ELLIPSIS
            End If
            If cell.Value <> shortName Then
                cell.Value = shortName
                If Len(originalName) - Len(shortName) > 50 Then
                    LogEvt "Formatting", lvlDETAIL, "Shortened device name by " & (Len(originalName) - Len(shortName)) & " chars", Left(shortName, 30) & "..."
                End If
            End If
            On Error Resume Next
            cell.AddComment Text:=originalName
            If Err.Number = 0 Then
                cell.Comment.Shape.TextFrame.AutoSize = True
            Else
                LogEvt "Formatting", lvlWARN, "Could not add comment to " & cell.Address & ": " & Err.Description
                Debug.Print Time & " - Warning: Could not add comment to cell " & cell.Address & ": " & Err.Description: Err.Clear
            End If
            On Error GoTo ShortNameErrorHandler
        End If
    Next cell
    Application.ScreenUpdating = True
    LogEvt "Formatting", lvlINFO, "Smart short names/comments processing complete. Processed " & devNameRange.Cells.Count & " device names."
    Exit Sub
ShortNameErrorHandler:
    Application.ScreenUpdating = True
    LogEvt "Formatting", lvlERROR, "Error applying smart short names/comments: " & Err.Description
    MsgBox "Error applying smart device names/comments: " & Err.Description, vbExclamation, "Short Name Error"
End Sub

Private Sub FreezeHeaderAndKeyCols(ws As Worksheet)
    ' Purpose: Freezes header row and key columns for better navigation.
    '          Freezes columns up to "Category" by default, or adjust as needed.
    On Error GoTo FreezeErrorHandler

    Dim tbl As ListObject
    Dim targetCol As ListColumn
    Dim freezeColIndex As Long
    Const COL_TO_FREEZE_AFTER As String = "Category" ' <<< Column to freeze panes AFTER
    Const FALLBACK_FREEZE_COL As Long = 4            ' Fallback if COL_TO_FREEZE_AFTER not found

    ' <<< CORRECTED: Check for table and set tbl object on separate lines >>>
    If ws.ListObjects.Count = 0 Then Exit Sub ' Exit if no table found on the worksheet
    Set tbl = ws.ListObjects(1)            ' Set the table object *after* the check

    LogEvt "Formatting", lvlDETAIL, "Applying freeze panes..."

    ' Find the column index to freeze after
    On Error Resume Next ' Handle if COL_TO_FREEZE_AFTER doesn't exist
    Set targetCol = Nothing ' Reset object
    Set targetCol = tbl.ListColumns(COL_TO_FREEZE_AFTER)
    On Error GoTo FreezeErrorHandler ' Restore error handler

    If targetCol Is Nothing Then
        freezeColIndex = FALLBACK_FREEZE_COL + 1 ' Use fallback column index + 1
        LogEvt "Formatting", lvlWARN, "Freeze column '" & COL_TO_FREEZE_AFTER & "' not found. Using fallback index: " & FALLBACK_FREEZE_COL
    Else
        freezeColIndex = targetCol.Index + 1 ' Use found column index + 1
    End If
    Set targetCol = Nothing ' Release object

    ' Ensure freeze index is valid
    If freezeColIndex < 2 Then freezeColIndex = 2 ' Minimum freeze is after column A (index 2)
    If freezeColIndex > ws.Columns.Count Then freezeColIndex = ws.Columns.Count ' Cannot freeze beyond last column

    Dim targetCell As Range
    Set targetCell = ws.Cells(tbl.HeaderRowRange.Row + 1, freezeColIndex) ' Cell below header, right of last frozen column

    ' Apply Freeze Panes
    ws.Activate ' Sheet must be active to set freeze panes on ActiveWindow
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False ' Unfreeze first if already frozen
    targetCell.Select ' Select the cell to freeze relative to
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select ' Select A1 for visual tidiness after freezing

    LogEvt "Formatting", lvlDETAIL, "Freeze Panes applied after column index " & freezeColIndex - 1 & " ('" & COL_TO_FREEZE_AFTER & "' or fallback)."
    Exit Sub ' Normal exit

FreezeErrorHandler:
    LogEvt "Formatting", lvlERROR, "Error applying freeze panes: " & Err.Description
    Debug.Print Time & " - Error applying freeze panes: " & Err.Description
    ' Non-critical error, don't stop execution, but maybe notify user
    ' MsgBox "Could not apply freeze panes: " & Err.Description, vbExclamation, "Freeze Panes Error"

End Sub

' ==========================================================================
' ===                       ARCHIVING FUNCTION                         ===
' ==========================================================================
Private Sub ArchiveMonth(wsDataSource As Worksheet, archiveSheetName As String)
    Dim wsArchive As Worksheet
    On Error GoTo ArchiveErrorHandler
    Application.DisplayAlerts = False
    LogEvt "Archive", lvlINFO, "Starting archive process for: " & archiveSheetName

    wsDataSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsArchive = ActiveSheet

    On Error Resume Next
    wsArchive.Name = Left(archiveSheetName, 31)
    If Err.Number <> 0 Then
        Dim fallbackName As String: fallbackName = "Archive_Error_" & Format(Now(), "yyyyMMdd_HHmmss")
        LogEvt "Archive", lvlWARN, "Rename to '" & archiveSheetName & "' failed. Using fallback: " & fallbackName
        wsArchive.Name = fallbackName: Err.Clear
    End If
    On Error GoTo ArchiveErrorHandler

    If wsArchive.UsedRange.Cells.CountLarge > 1 Then
        wsArchive.UsedRange.Value = wsArchive.UsedRange.Value
        LogEvt "Archive", lvlDETAIL, "Converted formulas to values on sheet: " & wsArchive.Name
    End If

    If wsArchive.ListObjects.Count > 0 Then
        On Error Resume Next
        wsArchive.ListObjects(1).Unlist
        LogEvt "Archive", lvlDETAIL, "Unlisted table on archive sheet: " & wsArchive.Name
        On Error GoTo ArchiveErrorHandler
    End If

    ' Optional: Clear Comments
    ' On Error Resume Next: wsArchive.Cells.ClearComments: LogEvt "Archive", lvlDETAIL, "Cleared comments from archive." : On Error GoTo ArchiveErrorHandler

    ' Optional: Protect Sheet
    ' On Error Resume Next
    ' wsArchive.Protect Password:="YourPassword", UserInterfaceOnly:=True
    ' LogEvt "Archive", lvlDETAIL, "Protected archive sheet: " & wsArchive.Name
    ' On Error GoTo ArchiveErrorHandler

    LogEvt "Archive", lvlINFO, "Successfully archived data to sheet: " & wsArchive.Name
    Application.DisplayAlerts = True
    Exit Sub

ArchiveErrorHandler:
    Dim errDesc As String: errDesc = Err.Description: Dim errNum As Long: errNum = Err.Number
    Application.DisplayAlerts = True
    LogEvt "Archive", lvlERROR, "Error during archiving for '" & archiveSheetName & "': " & errDesc & " (#" & errNum & ")"
    MsgBox "Error during archiving process for sheet '" & archiveSheetName & "': " & vbCrLf & errDesc, vbCritical, "Archive Error"
    If Not wsArchive Is Nothing Then
        If wsArchive.Name <> wsDataSource.Name Then
            On Error Resume Next
            wsArchive.Delete
            On Error GoTo 0
            LogEvt "Archive", lvlWARN, "Attempted delete of partial archive sheet due to error."
        End If
    End If
End Sub

' ==========================================================================
' ===                   HELPER & UTILITY FUNCTIONS                     ===
' ==========================================================================
Private Function GetWorksheets(ByRef wsData As Worksheet, ByRef wsWeights As Worksheet, ByRef wsCache As Worksheet) As Boolean
    Dim success As Boolean: success = True
    Const PROC_NAME As String = "GetWorksheets"
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Data sheet '" & DATA_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDETAIL, "Found wsData: " & wsData.Name
    Set wsWeights = ThisWorkbook.Sheets(WEIGHTS_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Weights sheet '" & WEIGHTS_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDETAIL, "Found wsWeights: " & wsWeights.Name
    Set wsCache = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    If Err.Number <> 0 Then LogEvt PROC_NAME, lvlERROR, "Cache sheet '" & CACHE_SHEET_NAME & "' not found.": success = False: Err.Clear Else LogEvt PROC_NAME, lvlDETAIL, "Found wsCache: " & wsCache.Name
    On Error GoTo 0
    If Not success Then
        MsgBox "Critical Error: One or more required worksheets could not be found." & vbCrLf & _
               "Ensure sheets named '" & DATA_SHEET_NAME & "', '" & WEIGHTS_SHEET_NAME & "', and '" & CACHE_SHEET_NAME & "' exist.", vbCritical, "Sheet Missing"
        Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing
    End If
    GetWorksheets = success
End Function

' <<< USING MINIMAL QUERYTABLE REFRESH VERSION >>>
Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Minimal version using QueryTable directly - avoids WorkbookConnection issues
    Dim qt As QueryTable
    On Error GoTo RefreshErrorHandler
    RefreshPowerQuery = False ' Default to False

    LogEvt "Refresh", lvlINFO, "Attempting QueryTable refresh for: " & targetTable.Name

    Set qt = targetTable.QueryTable
    If qt Is Nothing Then
        LogEvt "Refresh", lvlERROR, "Could not find QueryTable associated with table '" & targetTable.Name & "'."
        MsgBox "Error: Could not find QueryTable associated with table '" & targetTable.Name & "'.", vbCritical, "Refresh Error"
        Exit Function
    End If

    ' Use QueryTable properties for refresh
    qt.BackgroundQuery = False  ' Ensure foreground refresh
    qt.Refresh                  ' Refresh synchronously (no argument needed defaults to sync for QT)

    RefreshPowerQuery = True ' If refresh completes without error
    LogEvt "Refresh", lvlINFO, "QueryTable refresh completed successfully for: " & targetTable.Name
    Exit Function

RefreshErrorHandler:
    RefreshPowerQuery = False
    LogEvt "Refresh", lvlERROR, "Error during QueryTable refresh for '" & targetTable.Name & "'. Error #" & Err.Number & ": " & Err.Description
    MsgBox "QueryTable refresh failed for table '" & targetTable.Name & "': " & vbCrLf & Err.Description, vbExclamation, "Refresh Error"
    ' Consider Err.Raise here if the calling function needs to know a specific error occurred
    ' Err.Raise Number:=vbObjectError + 515, Source:="RefreshPowerQuery", Description:="QueryTable.Refresh failed."
End Function

Public Function IsMaintainerUser() As Boolean
    On Error Resume Next
    IsMaintainerUser = (LCase(Environ("USERNAME")) = LCase(MAINTAINER_USERNAME))
    If Err.Number <> 0 Then LogEvt "Util", lvlERROR, "Error checking MAINTAINER_USERNAME: " & Err.Description: IsMaintainerUser = False
    On Error GoTo 0
End Function

Private Function GetAPIKey() As String
    Dim fso As Object, ts As Object, keyPath As String, WshShell As Object, fileContent As String: fileContent = ""
    On Error GoTo KeyError
    keyPath = API_KEY_FILE_PATH
    Set WshShell = CreateObject("WScript.Shell")
    keyPath = WshShell.ExpandEnvironmentStrings(keyPath)
    Set WshShell = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(keyPath) Then
        Set ts = fso.OpenTextFile(keyPath, 1)
        If Not ts.AtEndOfStream Then fileContent = ts.ReadAll
        ts.Close
         LogEvt "APIKey", lvlDETAIL, "API Key read successfully from: " & keyPath ' Changed level
    Else
         LogEvt "APIKey", lvlWARN, "API Key file not found at: " & keyPath
        Debug.Print Time & " - WARNING: API Key file not found at specified path: " & keyPath
    End If
    GoTo KeyExit
KeyError:
     LogEvt "APIKey", lvlERROR, "Error reading API Key from '" & keyPath & "': " & Err.Description
    Debug.Print Time & " - ERROR reading API Key from '" & keyPath & "': " & Err.Description
KeyExit:
    GetAPIKey = Trim(fileContent)
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    On Error GoTo 0
End Function

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
    Set ws = Nothing
End Function

Private Function SafeGetString(arr As Variant, r As Long, ByVal cols As Object, colName As String) As String
    On Error Resume Next: SafeGetString = Trim(CStr(arr(r, cols(colName)))): If Err.Number <> 0 Then SafeGetString = "": Err.Clear
End Function
Private Function SafeGetVariant(arr As Variant, r As Long, ByVal cols As Object, colName As String) As Variant
    On Error Resume Next: SafeGetVariant = arr(r, cols(colName)): If Err.Number <> 0 Then SafeGetVariant = Null: Err.Clear
End Function

Private Function JsonEscape(strInput As String) As String
    strInput = Replace(strInput, "\", "\\")
    strInput = Replace(strInput, """", "\""")
    strInput = Replace(strInput, vbCrLf, "\n")
    strInput = Replace(strInput, vbCr, "\n")
    strInput = Replace(strInput, vbLf, "\n")
    strInput = Replace(strInput, vbTab, "\t")
    JsonEscape = strInput
End Function
Private Function JsonUnescape(strInput As String) As String
    strInput = Replace(strInput, "\""", """")
    strInput = Replace(strInput, "\\", "\")
    strInput = Replace(strInput, "\n", vbCrLf)
    strInput = Replace(strInput, "\t", vbTab)
    JsonUnescape = strInput
End Function

' --- Placeholder for Logger Module Calls (If Used) ---
Private Sub LogEvt(eventCode As String, eventLevel As Integer, eventDesc As String, Optional eventDetail As String = "")
    Dim logLevel As eLogLevel
    Select Case eventLevel
        Case 1: logLevel = lvlINFO
        Case 2: logLevel = lvlDETAIL
        Case 3: logLevel = lvlWARN
        Case 4, 5: logLevel = lvlERROR ' Map FATAL (5) to ERROR (4) if using integer levels
        Case Else: logLevel = lvlINFO
    End Select
    Dim levelString As String
    Select Case logLevel
        Case lvlINFO: levelString = "INFO"
        Case lvlDETAIL: levelString = "DETAIL"
        Case lvlWARN: levelString = "WARN"
        Case lvlERROR: levelString = "ERROR"
        Case Else: levelString = "UNKNOWN"
    End Select
    Debug.Print Time & " [" & levelString & "] " & _
                eventCode & ": " & eventDesc & IIf(eventDetail <> "", " | " & eventDetail, "")
    ' Forward to the unified logging system if it exists
    On Error Resume Next ' Prevent error if mod_Logger doesn't exist or LogEvt fails
    mod_Logger.LogEvt eventCode, logLevel, eventDesc, eventDetail
    On Error GoTo 0
End Sub

Private Sub FlushLogBuf()
    Debug.Print Time & " [INFO] LogFlush: Flushing log buffer to sheet"
    ' Forward to the unified logging system if it exists
    On Error Resume Next ' Prevent error if mod_Logger doesn't exist or FlushLogBuf fails
    mod_Logger.FlushLogBuf
    On Error GoTo 0
End Sub

' ==========================================================================
' ===                        END OF MODULE                               ===
' ==========================================================================


