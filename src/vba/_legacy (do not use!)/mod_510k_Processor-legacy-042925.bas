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
' --- Requires companion module: mod_Logger                              ---
' --- Requires companion module: mod_DebugTraceHelpers                   ---
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
Public Const DATA_SHEET_NAME As String = "CurrentMonthData"  ' Sheet where Power Query loads data (Public for ThisWorkbook)
Private Const WEIGHTS_SHEET_NAME As String = "Weights"        ' Sheet containing weight/keyword tables
Private Const CACHE_SHEET_NAME As String = "CompanyCache"      ' Sheet for persistent company recap cache
Private Const LOG_SHEET_NAME As String = "RunLog"             ' Defined here for reference, used by mod_Logger
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
Public Const VERSION_INFO As String = "v1.9 - Split Code Gen" ' Simple version tracking (Public for Logger)
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

' --- Module-Level Object for Regular Expressions (Late Binding) ---
Private regex As Object ' For CheckKeywords

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
    Dim originalCalcBeforeSave As Boolean ' For performance tweak
    Dim originalCalcState As XlCalculation ' For performance tweak

    ' --- Error Handling Setup ---
    On Error GoTo ProcessErrorHandler

    ' --- Initial Setup & Screen Handling ---
    Application.ScreenUpdating = False
    If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False ' Ensure off at start
    originalCalcState = Application.Calculation ' Store initial calc state
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = "Initializing 510(k) processing..."

    ' --- Initialize Logging ---
    ' Uses lg... constants from mod_Logger's Public Enum LogLevel
    ' Uses lvl... constants from mod_DebugTraceHelpers' Public Enum eTraceLvl
    LogEvt "ProcessStart", lgINFO, "ProcessMonthly510k Started", "Version=" & VERSION_INFO
    TraceEvt lvlINFO, "ProcessMonthly510k", "Process Start", "Version=" & VERSION_INFO ' Use enum

    ' --- Get Worksheet Objects Safely ---
    If Not GetWorksheets(wsData, wsWeights, wsCache) Then GoTo CleanExit ' Exit handled by EnsureUIOn

    ' --- Ensure Data Table Exists (Guard Rail) ---
    On Error Resume Next ' Check for table existence
    Set tblData = wsData.ListObjects(1) ' Try to get the first table
    On Error GoTo ProcessErrorHandler ' Restore error handler

    If tblData Is Nothing Then
        Dim rng As Range
        On Error Resume Next ' Handle errors during table creation
        Set rng = wsData.Range("A1").CurrentRegion ' Assume data starts at A1
        If Not rng Is Nothing And rng.Cells.Count > 1 Then ' Check if there's data
            Set tblData = wsData.ListObjects.Add(SourceType:=xlSrcRange, _
                                                 Source:=rng, _
                                                 XlListObjectHasHeaders:=xlYes)
            If Not tblData Is Nothing Then
                tblData.Name = "pgGet510kData_" & Format(Now, "yyyymmddhhmmss") ' Give unique name
                LogEvt "DataTable", lgWARN, "Table was missing – recreated from current region as '" & tblData.Name & "'."
                TraceEvt lvlWARN, "ProcessMonthly510k", "Data table missing, recreated as '" & tblData.Name & "'"
            Else
                LogEvt "DataTable", lgERROR, "Table was missing and failed to recreate from current region."
                TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, failed to recreate."
                GoTo ProcessErrorHandler ' Cannot proceed without a table
            End If
        Else
             LogEvt "DataTable", lgERROR, "Table was missing and no data found in CurrentRegion of A1 to recreate it."
             TraceEvt lvlERROR, "ProcessMonthly510k", "Data table missing, no data in A1 CurrentRegion."
             GoTo ProcessErrorHandler ' Cannot proceed without a table
        End If
        On Error GoTo ProcessErrorHandler ' Restore error handler
    End If
    ' --- End Table Guard Rail ---

    ' --- Determine Target Month & Check Guard Conditions ---
    startMonth = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
    targetMonthName = Format$(startMonth, "MMM-yyyy")
    archiveSheetName = targetMonthName
    mustArchive = Not SheetExists(archiveSheetName)
    proceed = mustArchive Or Day(Date) <= 5 Or IsMaintainerUser()

    LogEvt "ArchiveCheck", IIf(proceed, lgINFO, lgWARN), _
           "Guard conditions: Archive needed=" & mustArchive & _
           ", Day of month=" & Day(Date) & ", Is maintainer=" & IsMaintainerUser() & _
           ", Will proceed=" & proceed
    TraceEvt IIf(proceed, lvlINFO, lvlWARN), "ProcessMonthly510k", "Guard Check", "Proceed=" & proceed & ", ArchiveNeeded=" & mustArchive & ", Day=" & Day(Date) & ", Maintainer=" & IsMaintainerUser()

    If Not proceed Then
        LogEvt "ProcessSkip", lgINFO, "Processing skipped: Archive exists, not day 1-5, not maintainer."
        TraceEvt lvlINFO, "ProcessMonthly510k", "Processing Skipped (Guard Conditions Met)"
        Application.StatusBar = "Month " & targetMonthName & " already archived. Refreshing current view only."
        ' Attempt refresh even if skipping full process
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
        If tblData Is Nothing Then
            LogEvt "Refresh", lgERROR, "Data table not found on " & DATA_SHEET_NAME & " during skipped run check."
            TraceEvt lvlERROR, "ProcessMonthly510k", "Data table not found during skipped run refresh check"
        Else
            If Not RefreshPowerQuery(tblData) Then
                LogEvt "Refresh", lgERROR, "PQ Refresh failed during skipped run check."
                TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh failed during skipped run check"
                 ' Decide if this is critical enough to stop even a refresh-only run
                 ' GoTo ProcessErrorHandler ' Option to make it critical
            End If
        End If
        Set tblData = Nothing
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    Application.StatusBar = "Processing for month: " & targetMonthName

    ' --- Get Data Table & Check for Data (Redundant check, but safe) ---
    If tblData Is Nothing Then
        On Error Resume Next: Set tblData = wsData.ListObjects(1): On Error GoTo ProcessErrorHandler
    End If
    If tblData Is Nothing Then
        LogEvt "DataTable", lgERROR, "Data table not found on " & DATA_SHEET_NAME
        TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost or not found before refresh"
        GoTo ProcessErrorHandler
    End If

    ' --- Refresh Power Query Data (using the table object) ---
    Application.StatusBar = "Refreshing FDA data from Power Query..."
    LogEvt "Refresh", lgINFO, "Attempting PQ refresh for table: " & tblData.Name
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query Start", "Table=" & tblData.Name
    If Not RefreshPowerQuery(tblData) Then ' Includes post-refresh lock
         LogEvt "Refresh", lgERROR, "PQ Refresh failed. Processing stopped." ' Already logged in func, add context
         TraceEvt lvlERROR, "ProcessMonthly510k", "PQ Refresh Failed - Halting Process"
         GoTo ProcessErrorHandler ' Stop on critical PQ error
    End If
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Refresh Power Query End"

    ' --- Re-check table and data after refresh (PQ might empty it) ---
    If tblData Is Nothing Then ' Should not happen if RefreshPowerQuery succeeded, but defensive check
         LogEvt "DataTable", lgERROR, "Data table object became Nothing after refresh."
         TraceEvt lvlERROR, "ProcessMonthly510k", "Data table object lost after refresh"
         GoTo ProcessErrorHandler
    End If
    If tblData.ListRows.Count = 0 Then
        LogEvt "DataTable", lgWARN, "No data returned by Power Query for " & targetMonthName & "."
        TraceEvt lvlWARN, "ProcessMonthly510k", "No data after PQ refresh", "Month=" & targetMonthName
        MsgBox "No data returned by Power Query for " & targetMonthName & ". Nothing to process.", vbInformation, "No Data"
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    recordCount = tblData.ListRows.Count
    LogEvt "DataTable", lgINFO, "Table contains " & recordCount & " rows post-refresh."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Data Rows Post-Refresh", "Count=" & recordCount

    ' --- Add/Verify Output Columns ---
    LogEvt "Columns", lgINFO, "Checking/Adding scoring output columns..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns Start"
    If Not AddScoreColumnsIfNeeded(tblData) Then GoTo ProcessErrorHandler
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Add/Verify Columns End"

    ' --- Map Column Headers to Indices ---
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns Start"
    Set colIndices = GetColumnIndices(tblData.HeaderRowRange) ' Now handles duplicates
    If colIndices Is Nothing Then GoTo ProcessErrorHandler
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Map Columns End", "MappedKeys=" & colIndices.Count

    ' --- Load Weights, Keywords, and Cache ---
    Application.StatusBar = "Loading scoring parameters and cache..."
    LogEvt "LoadParams", lgINFO, "Loading weights, keywords, and cache..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters Start"
    If Not LoadWeightsAndKeywords(wsWeights) Then GoTo ProcessErrorHandler
    Call LoadCompanyCache(wsCache)
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Load Parameters End"

    ' --- Read Data into Array for Fast Processing ---
    Application.StatusBar = "Reading data into memory (" & recordCount & " rows)..."
    LogEvt "ReadData", lgINFO, "Reading data into array..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Read Data to Array Start", "Rows=" & recordCount
    ' Ensure dataArr is correctly dimensioned, especially for single row case
    If recordCount = 1 Then
        ' Handle single row specifically to ensure 2D array
        Dim singleRowData As Variant
        singleRowData = tblData.DataBodyRange.Value ' Read value first
        ' Check if it's already a 2D array (1 row, N columns)
        If IsArray(singleRowData) Then
             If UBound(singleRowData, 1) = 1 And UBound(singleRowData, 2) > 0 Then
                 dataArr = singleRowData ' It's already 2D(1, N)
             Else
                 ' It's an array but not 2D(1, N), needs manual creation
                 ReDim dataArr(1 To 1, 1 To tblData.ListColumns.Count)
                 Dim j As Long
                 Dim tempArr As Variant: tempArr = tblData.DataBodyRange.Value2 ' Use Value2 for raw data
                 For j = 1 To tblData.ListColumns.Count
                     On Error Resume Next ' Handle potential error reading individual cell
                     dataArr(1, j) = tempArr(j) ' Assumes tempArr is 1D for single row
                     If Err.Number <> 0 Then dataArr(1, j) = CVErr(xlErrNA): Err.Clear ' Assign error if read fails
                     On Error GoTo ProcessErrorHandler ' Restore handler
                 Next j
                 LogEvt "ReadData", lgDETAIL, "Manually created 2D array for single row from 1D array."
                 TraceEvt lvlDET, "ProcessMonthly510k", "Created 2D array for single row", "Source=1D Array"
             End If
        Else
             ' Single cell value, needs manual 2D array creation
             ReDim dataArr(1 To 1, 1 To tblData.ListColumns.Count)
             If tblData.ListColumns.Count = 1 Then
                 dataArr(1, 1) = singleRowData ' Directly assign single value
             Else
                 ' This case is less likely (single row result not being an array when >1 col), handle defensively
                  LogEvt "ReadData", lgWARN, "Single row read returned non-array despite multiple columns. Reading cell-by-cell."
                  TraceEvt lvlWARN, "ProcessMonthly510k", "Single row read as non-array", "Columns=" & tblData.ListColumns.Count
                  Dim k As Long
                  For k = 1 To tblData.ListColumns.Count
                    On Error Resume Next
                    dataArr(1, k) = tblData.DataBodyRange.Cells(1, k).Value2
                    If Err.Number <> 0 Then dataArr(1, k) = CVErr(xlErrNA): Err.Clear
                    On Error GoTo ProcessErrorHandler
                  Next k
             End If
             LogEvt "ReadData", lgDETAIL, "Manually created 2D array for single row from single value."
             TraceEvt lvlDET, "ProcessMonthly510k", "Created 2D array for single row", "Source=Single Value"
        End If
    ElseIf recordCount > 1 Then
        dataArr = tblData.DataBodyRange.Value2 ' Use Value2 for potentially faster read of raw data
    Else
        LogEvt "ReadData", lgWARN, "Attempted to read data array when recordCount is 0."
        TraceEvt lvlWARN, "ProcessMonthly510k", "Skipped reading data to array", "Reason=Zero Rows"
        GoTo CleanExit ' Exit handled by EnsureUIOn
    End If
    LogEvt "ReadData", lgINFO, "Read " & recordCount & " records into array (Ensured 2D)."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Read Data to Array End", "ArrayDims=" & LBound(dataArr, 1) & "-" & UBound(dataArr, 1) & ", " & LBound(dataArr, 2) & "-" & UBound(dataArr, 2)

    ' --- Main Processing Loop (Starts in Part 2) ---
    ' ... Code continues in Part 2 ...

' ==========================================================================
' ===                HELPER FUNCTIONS (Called by Part 1)               ===
' ==========================================================================

Private Function GetWorksheets(ByRef wsData As Worksheet, ByRef wsWeights As Worksheet, ByRef wsCache As Worksheet) As Boolean
    ' Purpose: Safely gets required worksheet objects by name.
    Dim success As Boolean: success = True
    Const PROC_NAME As String = "GetWorksheets"
    TraceEvt lvlDET, PROC_NAME, "Getting required worksheets..."

    On Error Resume Next ' Check each sheet individually
    Set wsData = Nothing: Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    If wsData Is Nothing Then LogEvt PROC_NAME, lgERROR, "Data sheet '" & DATA_SHEET_NAME & "' not found.": success = False: TraceEvt lvlERROR, PROC_NAME, "Sheet Not Found", "Name=" & DATA_SHEET_NAME Else TraceEvt lvlDET, PROC_NAME, "Found wsData", wsData.Name

    Set wsWeights = Nothing: Set wsWeights = ThisWorkbook.Sheets(WEIGHTS_SHEET_NAME)
    If wsWeights Is Nothing Then LogEvt PROC_NAME, lgERROR, "Weights sheet '" & WEIGHTS_SHEET_NAME & "' not found.": success = False: TraceEvt lvlERROR, PROC_NAME, "Sheet Not Found", "Name=" & WEIGHTS_SHEET_NAME Else TraceEvt lvlDET, PROC_NAME, "Found wsWeights", wsWeights.Name

    Set wsCache = Nothing: Set wsCache = ThisWorkbook.Sheets(CACHE_SHEET_NAME)
    If wsCache Is Nothing Then LogEvt PROC_NAME, lgERROR, "Cache sheet '" & CACHE_SHEET_NAME & "' not found.": success = False: TraceEvt lvlERROR, PROC_NAME, "Sheet Not Found", "Name=" & CACHE_SHEET_NAME Else TraceEvt lvlDET, PROC_NAME, "Found wsCache", wsCache.Name

    On Error GoTo 0 ' Restore default error handling

    If Not success Then
        MsgBox "Critical Error: One or more required worksheets could not be found." & vbCrLf & _
               "Ensure sheets named '" & DATA_SHEET_NAME & "', '" & WEIGHTS_SHEET_NAME & "', and '" & CACHE_SHEET_NAME & "' exist.", vbCritical, "Sheet Missing"
        Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing
        TraceEvt lvlERROR, PROC_NAME, "GetWorksheets failed."
    End If
    GetWorksheets = success
End Function

Public Function RefreshPowerQuery(targetTable As ListObject) As Boolean
    ' Purpose: Refreshes the Power Query associated with the target table using QueryTable object.
    '          Includes disabling background refresh post-query.
    Dim qt As QueryTable
    Const PROC_NAME As String = "RefreshPowerQuery"
    RefreshPowerQuery = False ' Default to False

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot refresh."
        TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    On Error GoTo RefreshErrorHandler
    LogEvt PROC_NAME, lgINFO, "Attempting QueryTable refresh for: " & targetTable.Name
    TraceEvt lvlINFO, PROC_NAME, "Start refresh", "Table='" & targetTable.Name & "'"

    Set qt = targetTable.QueryTable
    If qt Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Could not find QueryTable associated with table '" & targetTable.Name & "'."
        TraceEvt lvlERROR, PROC_NAME, "QueryTable object is Nothing", "Table='" & targetTable.Name & "'"
        MsgBox "Error: Could not find QueryTable associated with table '" & targetTable.Name & "'.", vbCritical, "Refresh Error"
        Exit Function ' Exit, cannot refresh
    End If

    ' Refresh synchronously
    qt.BackgroundQuery = False
    qt.Refresh

    ' --- Lock refresh settings post-query (per review suggestion) ---
    On Error Resume Next ' Best effort to disable these
    qt.BackgroundQuery = False
    qt.EnableRefresh = False
    If Err.Number <> 0 Then
         LogEvt PROC_NAME, lgWARN, "Could not disable BackgroundQuery/EnableRefresh after refresh for table '" & targetTable.Name & "'. Error: " & Err.Description
         TraceEvt lvlWARN, PROC_NAME, "Failed to set BackgroundQuery=False / EnableRefresh=False", "Table='" & targetTable.Name & "', Err=" & Err.Description
         Err.Clear
    Else
         LogEvt PROC_NAME, lgDETAIL, "Set BackgroundQuery=False and EnableRefresh=False post-refresh for table '" & targetTable.Name & "'."
         TraceEvt lvlDET, PROC_NAME, "Set BackgroundQuery=False / EnableRefresh=False post-refresh", "Table='" & targetTable.Name & "'"
    End If
    On Error GoTo RefreshErrorHandler ' Restore main handler for this sub
    ' --- End Lock ---

    RefreshPowerQuery = True ' If refresh completes without error
    LogEvt PROC_NAME, lgINFO, "QueryTable refresh completed successfully for: " & targetTable.Name
    TraceEvt lvlINFO, PROC_NAME, "Refresh successful", "Table='" & targetTable.Name & "'"
    Exit Function

RefreshErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    RefreshPowerQuery = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error during QueryTable refresh for '" & targetTable.Name & "'. Error #" & errNum & ": " & errDesc
    TraceEvt lvlERROR, PROC_NAME, "Error during QueryTable refresh", "Table='" & targetTable.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "QueryTable refresh failed for table '" & targetTable.Name & "': " & vbCrLf & errDesc, vbExclamation, "Refresh Error"
    ' Exit Function ' Exit implicitly after error handler
End Function

Private Function SheetExists(sheetName As String) As Boolean
    ' Purpose: Checks if a sheet with the given name exists in the workbook.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
    Set ws = Nothing ' Clean up
End Function

Public Function IsMaintainerUser() As Boolean
    ' Purpose: Checks if the current Windows user matches the defined maintainer username.
    On Error Resume Next ' Handle potential error reading environment variable
    IsMaintainerUser = (LCase(Environ$("USERNAME")) = LCase(MAINTAINER_USERNAME))
    If Err.Number <> 0 Then
         ' Use Debug.Print here as Logger might not be initialized yet or could cause recursion
         Debug.Print Time & " - Util - ERROR checking MAINTAINER_USERNAME (Environ User): " & Err.Description
         TraceEvt lvlERROR, "IsMaintainerUser", "Error reading Environ USERNAME", "Err=" & Err.Number & " - " & Err.Description
         IsMaintainerUser = False ' Default to false on error
    End If
    On Error GoTo 0
End Function

Private Function AddScoreColumnsIfNeeded(tblData As ListObject) As Boolean
    ' Purpose: Ensures all required scoring/output columns exist in the table, adding them if necessary.
    Dim requiredCols As Variant, colName As Variant, col As ListColumn, addedCol As Boolean: addedCol = False
    Dim currentHeaders As Object: Set currentHeaders = CreateObject("Scripting.Dictionary")
    currentHeaders.CompareMode = vbTextCompare ' Case-insensitive check for existence
    Dim hCell As Range
    Const PROC_NAME As String = "AddScoreColumnsIfNeeded"

    On Error GoTo AddColErrorHandler

    ' Build set of existing headers for quick lookup
    For Each hCell In tblData.HeaderRowRange.Cells
        Dim hName As String: hName = Trim(hCell.Value)
        If Len(hName) > 0 And Not currentHeaders.Exists(hName) Then
            currentHeaders.Add hName, True
        End If
    Next hCell

    requiredCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc", "Final_Score", "Score_Percent", "Category", "CompanyRecap")
    LogEvt PROC_NAME, lgDETAIL, "Verifying existence of calculated columns..."
    TraceEvt lvlDET, PROC_NAME, "Verifying calculated columns", "Table=" & tblData.Name & ", RequiredCount=" & UBound(requiredCols) + 1

    ' Loop through required columns and add if missing
    For Each colName In requiredCols
        If Not currentHeaders.Exists(colName) Then
            On Error Resume Next ' Handle error during Add operation specifically
            Set col = tblData.ListColumns.Add()
            Dim addErrNum As Long: addErrNum = Err.Number
            Dim addErrDesc As String: addErrDesc = Err.Description
            On Error GoTo AddColErrorHandler ' Restore main handler

            If addErrNum <> 0 Then
                LogEvt PROC_NAME, lgERROR, "Failed to add column '" & colName & "'. Error: " & addErrDesc
                TraceEvt lvlERROR, PROC_NAME, "Failed to add column", "Column=" & colName & ", Err=" & addErrNum & " - " & addErrDesc
                GoTo AddColErrorHandler ' Treat failure to add column as critical
            End If

            col.Name = colName ' Set the name *after* adding
            addedCol = True
            currentHeaders.Add colName, True ' Add to our set of headers
            LogEvt PROC_NAME, lgINFO, "Added missing column: " & colName
            TraceEvt lvlINFO, PROC_NAME, "Added missing column", "Column=" & colName
        End If
        Set col = Nothing ' Reset for next iteration
    Next colName

    AddScoreColumnsIfNeeded = True
    Exit Function

AddColErrorHandler:
     Dim errDesc As String: errDesc = Err.Description ' Capture error
     LogEvt PROC_NAME, lgERROR, "Error checking/adding columns to table '" & tblData.Name & "': " & errDesc
     TraceEvt lvlERROR, PROC_NAME, "Error checking/adding columns", "Table=" & tblData.Name & ", Err=" & Err.Number & " - " & errDesc
    MsgBox "Error verifying or adding required columns to table '" & tblData.Name & "': " & vbCrLf & errDesc, vbCritical, "Column Setup Error"
    AddScoreColumnsIfNeeded = False
End Function

Private Function GetColumnIndices(headerRange As Range) As Object ' Scripting.Dictionary or Nothing
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

Private Function LoadWeightsAndKeywords(wsWeights As Worksheet) As Boolean
    ' Purpose: Loads all weight and keyword tables from the specified sheet into memory.
    Const PROC_NAME As String = "LoadWeightsAndKeywords"
    Dim success As Boolean: success = True ' Assume success unless critical load fails
    On Error GoTo LoadErrorHandler ' General handler for non-critical table load issues

    LogEvt PROC_NAME, lgINFO, "Attempting to load weights and keywords from sheet: " & wsWeights.Name
    TraceEvt lvlINFO, PROC_NAME, "Start loading", "Sheet='" & wsWeights.Name & "'"

    ' Load each table, log/trace success or failure
    Set dictACWeights = LoadTableToDict(wsWeights, "tblACWeights")
    TraceEvt IIf(dictACWeights Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded AC Weights", IIf(dictACWeights Is Nothing, "FAILED", "Count=" & dictACWeights.Count)
    Set dictSTWeights = LoadTableToDict(wsWeights, "tblSTWeights")
    TraceEvt IIf(dictSTWeights Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded ST Weights", IIf(dictSTWeights Is Nothing, "FAILED", "Count=" & dictSTWeights.Count)
    Set dictPCWeights = LoadTableToDict(wsWeights, "tblPCWeights") ' Optional
    TraceEvt IIf(dictPCWeights Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded PC Weights (Optional)", IIf(dictPCWeights Is Nothing, "FAILED/MISSING", "Count=" & dictPCWeights.Count)
    Set highValKeywordsList = LoadTableToList(wsWeights, "tblKeywords")
    TraceEvt IIf(highValKeywordsList Is Nothing, lvlERROR, lvlDET), PROC_NAME, "Loaded HighVal Keywords", IIf(highValKeywordsList Is Nothing, "FAILED", "Count=" & highValKeywordsList.Count)
    Set nfCosmeticKeywordsList = LoadTableToList(wsWeights, "tblNFCosmeticKeywords") ' Optional
    TraceEvt IIf(nfCosmeticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Cosmetic Keywords (Optional)", IIf(nfCosmeticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & nfCosmeticKeywordsList.Count)
    Set nfDiagnosticKeywordsList = LoadTableToList(wsWeights, "tblNFDiagnosticKeywords") ' Optional
    TraceEvt IIf(nfDiagnosticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Diagnostic Keywords (Optional)", IIf(nfDiagnosticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & nfDiagnosticKeywordsList.Count)
    Set therapeuticKeywordsList = LoadTableToList(wsWeights, "tblTherapeuticKeywords") ' Optional
    TraceEvt IIf(therapeuticKeywordsList Is Nothing, lvlWARN, lvlDET), PROC_NAME, "Loaded Therapeutic Keywords (Optional)", IIf(therapeuticKeywordsList Is Nothing, "FAILED/MISSING", "Count=" & therapeuticKeywordsList.Count)

    ' --- Critical Check: Ensure essential tables were loaded ---
    If dictACWeights Is Nothing Or dictSTWeights Is Nothing Or highValKeywordsList Is Nothing Then
         LogEvt PROC_NAME, lgERROR, "Critical failure: Could not load AC/ST weights or HighValue Keywords."
         TraceEvt lvlERROR, PROC_NAME, "CRITICAL FAILURE: Missing essential Weights/Keywords", "AC=" & IIf(dictACWeights Is Nothing, "FAIL", "OK") & ", ST=" & IIf(dictSTWeights Is Nothing, "FAIL", "OK") & ", KW=" & IIf(highValKeywordsList Is Nothing, "FAIL", "OK")
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

    TraceEvt lvlINFO, PROC_NAME, "Loading complete", "Success=True"
    LoadWeightsAndKeywords = True ' Indicate success
    Exit Function

LoadErrorHandler: ' Handles non-critical errors (e.g., optional table missing)
    Dim errDesc As String: errDesc = Err.Description
     LogEvt PROC_NAME, lgWARN, "Non-critical error loading one or more weight/keyword tables: " & errDesc & ". Defaults may be used.", "Sheet=" & wsWeights.Name
     TraceEvt lvlWARN, PROC_NAME, "Non-critical load error occurred", "Sheet='" & wsWeights.Name & "', Err=" & Err.Number & " - " & errDesc
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
     LoadWeightsAndKeywords = True ' Still return True for non-critical errors, allowing defaults
     Exit Function ' Need to explicitly exit after handling non-critical error

LoadErrorCritical: ' Handles failure to load essential tables
    MsgBox "Critical Error: Could not load essential AC/ST weights or HighValue Keywords from sheet '" & wsWeights.Name & "'. Processing cannot continue.", vbCritical, "Load Failure"
    TraceEvt lvlERROR, PROC_NAME, "Exiting due to critical load failure."
    ' Clean up any potentially partially loaded objects
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    LoadWeightsAndKeywords = False ' Indicate critical failure
End Function

Private Function LoadTableToDict(ws As Worksheet, tableName As String) As Object ' Scripting.Dictionary or Nothing
    ' Purpose: Loads a 2-column table into a Dictionary. Returns Nothing on error.
    Dim dict As Object ' Late bound Scripting.Dictionary
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, key As String, val As Variant
    Const PROC_NAME As String = "LoadTableToDict"

    On Error GoTo LoadDictError

    Set tbl = ws.ListObjects(tableName) ' This will error if table doesn't exist

    If tbl.ListRows.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "Table '" & tableName & "' is empty. Returning empty dictionary.", "Sheet=" & ws.Name
        TraceEvt lvlINFO, PROC_NAME, "Table empty", "Table=" & tableName
        Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare ' Return empty dict
        Set LoadTableToDict = dict
        Exit Function
    End If

    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then ' Should not happen if rows > 0, but check
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has rows but no DataBodyRange? Returning Nothing.", "Sheet=" & ws.Name
        TraceEvt lvlWARN, PROC_NAME, "No DataBodyRange despite rows>0", "Table=" & tableName
        Set LoadTableToDict = Nothing ' Indicate failure
        Exit Function
    End If

    If dataRange.Columns.Count < 2 Then
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has less than 2 columns. Cannot create dictionary. Returning Nothing.", "Sheet=" & ws.Name
        TraceEvt lvlWARN, PROC_NAME, "Table has < 2 columns", "Table=" & tableName
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
                  TraceEvt lvlWARN, PROC_NAME, "Unexpected array structure", "Table=" & tableName
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
    TraceEvt lvlWARN, PROC_NAME, "Error loading table to Dict", "Table=" & tableName & ", Err=" & Err.Number & " - " & Err.Description
    Set LoadTableToDict = Nothing ' Return Nothing on error
    ' No Resume here, exit handled by returning Nothing
End Function

Private Function LoadTableToList(ws As Worksheet, tableName As String) As Collection ' Returns Collection or Nothing
    ' Purpose: Loads the first column of a table into a Collection. Returns Nothing on error.
    Dim coll As Collection ' Use New Collection if early bound, otherwise create later
    Dim tbl As ListObject, dataRange As Range, dataArr As Variant, i As Long, item As String
    Const PROC_NAME As String = "LoadTableToList"

    On Error GoTo LoadListError

    Set tbl = ws.ListObjects(tableName) ' Errors if table doesn't exist

    If tbl.ListRows.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "Table '" & tableName & "' is empty. Returning empty collection.", "Sheet=" & ws.Name
        TraceEvt lvlINFO, PROC_NAME, "Table empty", "Table=" & tableName
        Set coll = New Collection ' Return new empty collection
        Set LoadTableToList = coll
        Exit Function
    End If

    Set dataRange = tbl.ListColumns(1).DataBodyRange ' Get first column data
    If dataRange Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "Table '" & tableName & "' has rows but no DataBodyRange in first column? Returning Nothing.", "Sheet=" & ws.Name
        TraceEvt lvlWARN, PROC_NAME, "No DataBodyRange in Col1", "Table=" & tableName
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
    TraceEvt lvlWARN, PROC_NAME, "Error loading table to List", "Table=" & tableName & ", Err=" & Err.Number & " - " & Err.Description
    Set LoadTableToList = Nothing ' Return Nothing on error
End Function

Private Sub LoadCompanyCache(wsCache As Worksheet)
    ' Purpose: Loads the persistent company cache from the sheet into memory.
    Dim lastRow As Long, i As Long, cacheData As Variant, loadedCount As Long
    Const PROC_NAME As String = "LoadCompanyCache"
    Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare
    TraceEvt lvlINFO, PROC_NAME, "Loading cache from sheet", "Sheet=" & wsCache.Name

    On Error GoTo CacheLoadError ' Use specific handler

    lastRow = wsCache.Cells(wsCache.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo ExitLoadCache ' No data rows

    ' Read CompanyName and RecapText only (Columns A and B)
    cacheData = wsCache.Range("A2:B" & lastRow).Value2

    If IsArray(cacheData) Then
        For i = 1 To UBound(cacheData, 1)
            Dim k As String: k = Trim(CStr(cacheData(i, 1)))
            Dim v As String: v = CStr(cacheData(i, 2))
            If Len(k) > 0 Then
                ' Overwrite if exists, add if new
                dictCache(k) = v
            End If
        Next i
    ElseIf lastRow = 2 Then ' Handle single data row case (Value2 might return single value if only 1 cell)
        ' Re-read explicitly as range values for single row if cacheData wasn't an array
        Dim kS As String: kS = Trim(CStr(wsCache.Range("A2").Value2))
        Dim vS As String: vS = CStr(wsCache.Range("B2").Value2)
        If Len(kS) > 0 Then dictCache(kS) = vS
    End If

ExitLoadCache:
    loadedCount = dictCache.Count
     LogEvt PROC_NAME, lgINFO, "Loaded " & loadedCount & " items into memory cache."
     TraceEvt lvlINFO, PROC_NAME, "Cache loading complete", "ItemsLoaded=" & loadedCount
    On Error GoTo 0 ' Ensure normal error handling restored
    Exit Sub

CacheLoadError:
     LogEvt PROC_NAME, lgERROR, "Error reading cache data from sheet: " & Err.Description
     TraceEvt lvlERROR, PROC_NAME, "Error reading cache data", "Sheet=" & wsCache.Name & ", Err=" & Err.Number & " - " & Err.Description
     Err.Clear ' Clear error before resuming
     Resume ExitLoadCache ' Go to cleanup/logging part
End Sub

Private Function SafeGetString(arr As Variant, r As Long, ByVal cols As Object, baseColName As String) As String
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

Private Function SafeGetVariant(arr As Variant, r As Long, ByVal cols As Object, baseColName As String) As Variant
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

Private Function SafeGetColIndex(colsDict As Object, baseColName As String) As Long
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

Private Function ColumnExistsInMap(dict As Object, baseColName As String) As Boolean
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

Private Function ConcatArrays(arr1 As Variant, arr2 As Variant) As Variant
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

    ' --- Main Processing Loop ---
    Application.StatusBar = "Calculating scores and fetching recaps (0% Complete)..."
    LogEvt "ScoreLoop", lgINFO, "Starting main processing loop for " & recordCount & " records."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Score Loop Start", "Rows=" & recordCount
    useOpenAI = IsMaintainerUser() ' Check if maintainer for OpenAI calls

    For i = 1 To recordCount
        ' Calculate Score
        scoreResult = Calculate510kScore(dataArr, i, colIndices)

        ' Get Company Recap (with Cache & Optional OpenAI)
        Dim companyName As String
        companyName = SafeGetString(dataArr, i, colIndices, "Applicant") ' Use safe getter
        If Len(companyName) > 0 Then
            currentRecap = GetCompanyRecap(companyName, useOpenAI)
        Else
            currentRecap = "Invalid Applicant Name"
            LogEvt "ScoreLoop", lgWARN, "Row " & i & ": Invalid/blank Applicant name."
            TraceEvt lvlWARN, "ScoreLoop", "Invalid Applicant Name", "Row=" & i
        End If

        ' Write results back to the memory array
        WriteResultsToArray dataArr, i, colIndices, scoreResult, currentRecap

        ' Update Status Bar periodically
        If i Mod 50 = 0 Or i = recordCount Then
            Application.StatusBar = "Calculating scores and fetching recaps (" & Format(i / recordCount, "0%") & " Complete)..."
            DoEvents ' Allow UI updates and prevent Excel from appearing frozen
        End If

        ' Log progress periodically
        If i Mod 100 = 0 Then ' Keep standard log for milestones
            LogEvt "ScoreLoop", lgDETAIL, "Processed " & i & " of " & recordCount & " records (" & Format(i / recordCount, "0%") & ")"
        End If
        If i Mod 50 = 0 Then ' Add more granular trace for spam level
             TraceEvt lvlSPAM, "ScoreLoop", "Processing row", "Row=" & i & "/" & recordCount ' Use enum
        End If
    Next i
    LogEvt "ScoreLoop", lgINFO, "Main processing loop complete."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Score Loop End"

    ' --- Write Processed Array Back to Sheet ---
    Application.StatusBar = "Writing results back to Excel sheet..."
    LogEvt "WriteBack", lgINFO, "Writing " & recordCount & " rows back to table '" & tblData.Name & "'."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Write Back Start", "Rows=" & recordCount

    ' <<< MODIFIED: Performance tweak for large write-back >>>
    originalCalcBeforeSave = Application.CalculateBeforeSave
    ' originalCalcState was already stored and set to manual
    Application.CalculateBeforeSave = False
    ' Ensure calculation is still manual just before write
    Application.Calculation = xlCalculationManual
    ' Optional: Application.CalculateFullRebuild ' If needed before write
    ' >>>

    On Error GoTo WriteBackError ' Specific error handler for the write operation
    tblData.DataBodyRange.Value = dataArr ' The actual write operation
    On Error GoTo ProcessErrorHandler ' Restore main handler

    ' <<< MODIFIED: Restore calculation settings >>>
    Application.CalculateBeforeSave = originalCalcBeforeSave
    ' Restore original state (which might have been auto)
    ' Application.Calculation = originalCalcState '<<< Let EnsureUIOn handle final restore >>>
    ' >>>

    LogEvt "WriteBack", lgINFO, "Array write complete."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Write Back End"

    ' --- Apply Number Formats ---
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Apply Number Formats Start"
    ApplyNumberFormats tblData
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Apply Number Formats End"

    ' --- Sort Table by DecisionDate ---
    Application.StatusBar = "Sorting data..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Sort Table Start"
    SortDataTable tblData, "DecisionDate", xlDescending
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Sort Table End"

    ' --- Save Updated Company Cache ---
    Application.StatusBar = "Saving company cache..."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Save Cache Start"
    If Not wsCache Is Nothing And Not dictCache Is Nothing Then
        If dictCache.Count > 0 Then
            LogEvt "SaveCache", lgINFO, "Saving " & dictCache.Count & " items to cache sheet '" & wsCache.Name & "'."
            Call SaveCompanyCache(wsCache)
        Else
            LogEvt "SaveCache", lgINFO, "In-memory cache is empty, skipping save to sheet."
            TraceEvt lvlINFO, "ProcessMonthly510k", "Skipped saving empty cache"
        End If
    Else
         LogEvt "SaveCache", lgWARN, "Cache sheet or dictionary object invalid, skipping save."
         TraceEvt lvlWARN, "ProcessMonthly510k", "Skipped saving cache", "Reason=Invalid Objects"
    End If
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Save Cache End"

    ' --- Final Layout, Formatting & Visual Polish (Applied before Archive) ---
    Application.StatusBar = "Applying final layout and formatting..."
    LogEvt "Formatting", lgINFO, "Applying final layout and formatting."
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Final Formatting Start"
    If Not tblData Is Nothing Then
        ' REVISED: Call ReorganizeColumns FIRST - using the robust copy/delete version
        Call ReorganizeColumns(tblData) ' Handles duplicates
        ' --- Optional: Assert Column Order Here ---
        ' Call AssertColumnOrder(tblData, Array("K_Number", "DecisionDate", ...)) ' Example
        ' ---
        Call FormatTableLook(wsData)    ' Custom formatting
        Call FormatCategoryColors(tblData)
        Call CreateShortNamesAndComments(tblData)
        ' Call FreezeHeaderAndKeyCols(wsData) ' <<< Still Disabled
        LogEvt "Formatting", lgINFO, "Final formatting applied."
        TraceEvt lvlINFO, "ProcessMonthly510k", "Final formatting applied successfully"
    Else
        LogEvt "Formatting", lgWARN, "Table object invalid before final formatting block."
        TraceEvt lvlWARN, "ProcessMonthly510k", "Skipped final formatting", "Reason=Invalid Table Object"
    End If
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Final Formatting End"
    ' --- End of Final Layout ---

    ' --- Archive Month (if needed) ---
    If proceed And mustArchive Then ' Check flags again before archiving
        Application.StatusBar = "Archiving month: " & targetMonthName & "..."
        TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archive Start", "ArchiveSheet=" & archiveSheetName
        If Not SheetExists(archiveSheetName) Then
            LogEvt "Archive", lgINFO, "Starting archive creation for " & targetMonthName & "."
            Call ArchiveMonth(wsData, archiveSheetName)
            TraceEvt lvlINFO, "ProcessMonthly510k", "Archive sheet created", "Sheet=" & archiveSheetName
        Else
            LogEvt "Archive", lgWARN, "Archive sheet '" & archiveSheetName & "' already exists - skipping creation."
            TraceEvt lvlWARN, "ProcessMonthly510k", "Skipped archive creation", "Reason=Sheet Exists"
        End If
        TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archive End"
    Else
         LogEvt "Archive", lgINFO, "Skipping archive step based on guard conditions."
         TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Archive Skipped", "Proceed=" & proceed & ", MustArchive=" & mustArchive
    End If


    ' --- Clean up duplicate connections created by sheet copy ---
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Cleanup Connections Start"
    Dim c As WorkbookConnection
    Dim baseConnectionName As String
    Dim originalConnection As WorkbookConnection ' To store ref if found
    Set originalConnection = Nothing
    On Error Resume Next
    Set originalConnection = ThisWorkbook.Connections("pgGet510kData")
    If originalConnection Is Nothing Then Set originalConnection = ThisWorkbook.Connections("Query - pgGet510kData")
    On Error GoTo ProcessErrorHandler ' Restore handler

    If Not originalConnection Is Nothing Then
        baseConnectionName = originalConnection.Name
        LogEvt "Cleanup", lgINFO, "Checking for duplicate connections based on found connection: '" & baseConnectionName & "'"
        TraceEvt lvlINFO, "ProcessMonthly510k", "Checking duplicate connections", "Base=" & baseConnectionName
        On Error Resume Next ' Ignore errors during loop/delete
        For Each c In ThisWorkbook.Connections
            If c.Name <> baseConnectionName And c.Name Like baseConnectionName & " (*" Then
                LogEvt "Cleanup", lgDETAIL, "Deleting duplicate connection: " & c.Name
                TraceEvt lvlDET, "ProcessMonthly510k", "Deleting duplicate connection", c.Name
                c.Delete
            End If
        Next c
        On Error GoTo ProcessErrorHandler ' Restore handler
    Else
         baseConnectionName = "pgGet510kData" ' Fallback base name
         LogEvt "Cleanup", lgWARN, "Could not find original PQ connection by typical names. Attempting cleanup based on: '" & baseConnectionName & "'"
         TraceEvt lvlWARN, "ProcessMonthly510k", "Original PQ Connection not found", "FallbackBase=" & baseConnectionName
         On Error Resume Next ' Ignore errors during loop/delete
         For Each c In ThisWorkbook.Connections
             If c.Name Like baseConnectionName & " (*" Or c.Name Like "Query - " & baseConnectionName & " (*" Then
                 LogEvt "Cleanup", lgDETAIL, "Deleting potential duplicate connection: " & c.Name
                 TraceEvt lvlDET, "ProcessMonthly510k", "Deleting potential duplicate connection", c.Name
                 c.Delete
             End If
         Next c
         On Error GoTo ProcessErrorHandler ' Restore handler
    End If
    Set c = Nothing
    Set originalConnection = Nothing
    TraceEvt lvlINFO, "ProcessMonthly510k", "Phase: Cleanup Connections End"
    ' --- End of duplicate connection cleanup ---

    ' --- Completion Message ---
    Dim endTime As Double: endTime = Timer
    Dim elapsed As String: elapsed = Format(endTime - startTime, "0.00")
    LogEvt "ProcessEnd", lgINFO, "Processing completed successfully.", "Records=" & recordCount & ", Elapsed=" & elapsed & "s"
    TraceEvt lvlINFO, "ProcessMonthly510k", "Process End", "Records=" & recordCount & ", Elapsed=" & elapsed & "s"
    Application.StatusBar = "Processing complete for " & targetMonthName & "."
    MsgBox "Monthly 510(k) data processed" & IIf(mustArchive, " and archived", "") & " for " & targetMonthName & "." & vbCrLf & vbCrLf & _
           "Processed " & recordCount & " records in " & elapsed & " seconds.", vbInformation, "Processing Complete"

CleanExit:
    LogEvt "Cleanup", lgINFO, "CleanExit reached. Releasing objects and restoring settings."
    TraceEvt lvlINFO, "ProcessMonthly510k", "CleanExit reached"
    ' Release objects
    Set dictACWeights = Nothing: Set dictSTWeights = Nothing: Set dictPCWeights = Nothing
    Set highValKeywordsList = Nothing: Set nfCosmeticKeywordsList = Nothing
    Set nfDiagnosticKeywordsList = Nothing: Set therapeuticKeywordsList = Nothing
    Set dictCache = Nothing: Set colIndices = Nothing
    Set wsData = Nothing: Set wsWeights = Nothing: Set wsCache = Nothing: Set wsLog = Nothing: Set tblData = Nothing
    Set regex = Nothing ' Release RegExp object

    ' Restore settings using the new failsafe routine
    Call EnsureUIOn ' Use central routine

    Debug.Print Time & " - ProcessMonthly510k Finished. Objects released."
    On Error Resume Next ' Ensure Flush doesn't trigger error handler again
    FlushLogBuf
    On Error GoTo 0
    Exit Sub

WriteBackError: ' Specific handler for array write-back
    LogEvt "WriteBack", lgERROR, "FATAL ERROR writing array back to sheet '" & tblData.Name & "': " & Err.Description
    TraceEvt lvlERROR, "ProcessMonthly510k", "FATAL ERROR writing array to sheet", "Table='" & tblData.Name & "', Err=" & Err.Number & " - " & Err.Description
    MsgBox "CRITICAL ERROR: Failed to write processed data back to the worksheet." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "The sheet may be in an inconsistent state. Please review the data and logs.", vbCritical, "Data Write Error"
    ' Restore calculation settings even after write error before jumping to main handler
    Application.CalculateBeforeSave = originalCalcBeforeSave
    ' Application.Calculation = originalCalcState ' Let EnsureUIOn handle final state
    Resume ProcessErrorHandler ' Jump to the main error handler for cleanup

ProcessErrorHandler:
      Dim errNum As Long: errNum = Err.Number
      Dim errDesc As String: errDesc = Err.Description
      Dim errLine As Long: errLine = Erl ' Get line number if available
      Dim errSource As String: errSource = Err.Source

      ' Avoid logging if error came FROM logger/tracer itself (basic check)
      If Not (errSource Like "*mod_Logger*" Or errSource Like "*mod_DebugTraceHelpers*") Then
           LogEvt "ProcessError", lgERROR, "Error #" & errNum & " at line " & errLine & " in " & errSource & ": " & errDesc
           TraceEvt lvlERROR, "ProcessErrorHandler", "Caught Error", "Err#" & errNum & ", Line=" & errLine & ", Source=" & errSource & ", Desc=" & errDesc
      End If
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
      Debug.Print Time & " - FATAL ERROR #" & errNum & " at line " & errLine & " in " & errSource & " (ProcessMonthly510k): " & errDesc
      Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"

      ' Show message box only once if multiple errors occur rapidly (though EnsureUIOn might redisplay it)
      Static errorMsgShown As Boolean
      If Not errorMsgShown Then
          MsgBox "A critical error occurred during 510(k) processing:" & vbCrLf & vbCrLf & _
                 "Error Number: " & errNum & vbCrLf & _
                 "Module: mod_510k_Processor" & vbCrLf & _
                 "Procedure: ProcessMonthly510k" & vbCrLf & _
                 "Source: " & errSource & vbCrLf & _
                 "Description: " & errDesc & vbCrLf & vbCrLf & _
                 "Attempting cleanup. Processing stopped.", vbCritical, "Processing Error"
          errorMsgShown = True ' Prevent repeated messages if error handler itself loops/errors
      End If

      ' Attempt to flush logs and restore UI settings
      On Error Resume Next ' Use Resume Next carefully in error handler
      FlushLogBuf
      Call EnsureUIOn ' Use central routine for cleanup
      errorMsgShown = False ' Reset flag after cleanup attempt
      On Error GoTo 0
      ' Exit Sub implicitly follows
End Sub


' ==========================================================================
' ===                CORE SCORING FUNCTION                         ===
' ==========================================================================
Private Function Calculate510kScore(dataArr As Variant, rowIdx As Long, ByVal cols As Object) As Variant
    ' Purpose: Calculates the 510(k) score based on various factors for a single record.
    ' Inputs:  dataArr - The 2D variant array holding all data.
    '          rowIdx - The current row number being processed in the array.
    '          cols - Dictionary mapping column names (including Name#Index for duplicates) to indices.
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

    ' --- 1. Extract Data Using Column Indices (Use SafeGetString/Variant which handle potential key issues) ---
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

    ' Processing Time Weight
    If IsNumeric(ProcTimeDays) Then
        Select Case CDbl(ProcTimeDays)
            Case Is > 172: PT_Wt = 0.65
            Case 162 To 172: PT_Wt = 0.6
            Case Else: PT_Wt = DEFAULT_PT_WEIGHT ' Includes < 162 and non-positive/invalid
        End Select
    Else: PT_Wt = DEFAULT_PT_WEIGHT ' Default if ProcTimeDays is not numeric
    End If

    ' Geographic Location Weight
    If Country = "US" Then GL_Wt = US_GL_WEIGHT Else GL_Wt = OTHER_GL_WEIGHT

    ' Keyword Weight (using RegExp function now)
    HasHighValueKW = CheckKeywords(combinedText, highValKeywordsList)
    If HasHighValueKW Then KW_Wt = HIGH_KW_WEIGHT Else KW_Wt = LOW_KW_WEIGHT

    ' --- 3. Negative Factors (NF) & Synergy Logic ---
    NF_Calc = 0: Synergy_Calc = 0
    IsCosmetic = CheckKeywords(combinedText, nfCosmeticKeywordsList)
    IsDiagnostic = CheckKeywords(combinedText, nfDiagnosticKeywordsList)
    HasTherapeuticMention = CheckKeywords(combinedText, therapeuticKeywordsList)

    ' Apply Negative Factors (Ensure Therapeutic overrides NF)
    If IsCosmetic And Not HasTherapeuticMention Then NF_Calc = NF_COSMETIC
    If IsDiagnostic And Not HasTherapeuticMention Then
        ' Additive logic: If both Cosmetic and Diagnostic (and not Therapeutic) apply both NFs
        If NF_Calc = 0 Then NF_Calc = NF_DIAGNOSTIC Else NF_Calc = NF_Calc + NF_DIAGNOSTIC
    End If

    ' Apply Synergy Bonus
    If (AC = "OR" Or AC = "NE") And HasHighValueKW Then Synergy_Calc = SYNERGY_BONUS

    ' --- 4. Final Score Calculation ---
    ' Ensure divisor matches the number of components being summed (adjust if logic changes)
    Final_Score_Raw = (AC_Wt + PC_Wt + KW_Wt + ST_Wt + PT_Wt + GL_Wt + NF_Calc + Synergy_Calc) / 6
    If Final_Score_Raw < 0 Then Final_Score_Raw = 0 ' Floor score at 0

    ' --- 5. Determine Category ---
    Select Case Final_Score_Raw
        Case Is > 0.6: Category = "High"
        Case 0.5 To 0.6: Category = "Moderate"
        Case 0.4 To 0.499999999999: Category = "Low" ' Explicit upper bound for Low
        Case Else: Category = "Almost None" ' Includes scores < 0.4 and exactly 0
    End Select

    ' --- 6. Return Results ---
    Calculate510kScore = Array(Final_Score_Raw, Category, AC_Wt, PC_Wt, KW_Wt, ST_Wt, PT_Wt, GL_Wt, NF_Calc, Synergy_Calc)
    Exit Function

ScoreErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    LogEvt "ScoreError", lgERROR, "Error scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc, "AC=" & AC & ", PC=" & PC ' Use lgERROR
    TraceEvt lvlERROR, "Calculate510kScore", "Error scoring row", "Row=" & rowIdx & ", K#=" & kNum & ", Err=" & Err.Number & " - " & errDesc
    Debug.Print Time & " - ERROR scoring row " & rowIdx & " (K#: " & kNum & "): " & errDesc
    Calculate510kScore = Array(0, "Error", 0, 0, 0, 0, 0, 0, 0, 0) ' Return error state array
End Function


' ==========================================================================
' ===                COMPANY RECAP & CACHING FUNCTIONS                   ===
' ==========================================================================
Private Function GetCompanyRecap(companyName As String, useOpenAI As Boolean) As String
    ' Purpose: Retrieves company recap, using memory cache, sheet cache, or optionally OpenAI.
    Dim finalRecap As String
    Const PROC_NAME As String = "GetCompanyRecap"

    ' Initialize cache dictionary if it's not already set
    If dictCache Is Nothing Then Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare

    ' Handle invalid input
    If Len(Trim(companyName)) = 0 Then
        LogEvt PROC_NAME, lgWARN, "Invalid (empty) company name passed.", "RowContext=Unknown" ' Use lgWARN
        TraceEvt lvlWARN, PROC_NAME, "Invalid company name passed", "Name=''"
        GetCompanyRecap = "Invalid Applicant Name"
        Exit Function
    End If

    ' 1. Check Memory Cache (Fastest)
    If dictCache.Exists(companyName) Then
        finalRecap = dictCache(companyName)
        LogEvt PROC_NAME, lgDETAIL, "Memory Cache HIT.", "Company=" & companyName ' Use lgDETAIL
        TraceEvt lvlDET, PROC_NAME, "Memory Cache Hit", "Company=" & companyName
    Else
        ' 2. Memory Cache MISS - Try OpenAI (if enabled) or use Default
        LogEvt PROC_NAME, lgDETAIL, "Memory Cache MISS.", "Company=" & companyName ' Use lgDETAIL
        TraceEvt lvlDET, PROC_NAME, "Memory Cache Miss", "Company=" & companyName

        finalRecap = DEFAULT_RECAP_TEXT ' Assume default unless OpenAI succeeds

        If useOpenAI Then
            Dim openAIResult As String
            LogEvt PROC_NAME, lgINFO, "Attempting OpenAI call.", "Company=" & companyName ' Use lgINFO
            TraceEvt lvlINFO, PROC_NAME, "Attempting OpenAI call", "Company=" & companyName
            openAIResult = GetCompanyRecapOpenAI(companyName) ' This function logs its own success/failure

            ' Update finalRecap only if OpenAI returns a valid, non-error result
            If openAIResult <> "" And Not LCase(openAIResult) Like "error:*" Then
                finalRecap = openAIResult
                 TraceEvt lvlINFO, PROC_NAME, "OpenAI SUCCESS, using result.", "Company=" & companyName
            Else
                 TraceEvt IIf(LCase(openAIResult) Like "error:*", lvlERROR, lvlWARN), PROC_NAME, "OpenAI Failed or Skipped, using default.", "Company=" & companyName & ", Result=" & openAIResult
            End If
        Else
             LogEvt PROC_NAME, lgINFO, "OpenAI call skipped (Not Maintainer or disabled).", "Company=" & companyName ' Use lgINFO
             TraceEvt lvlINFO, PROC_NAME, "OpenAI call skipped", "Company=" & companyName
        End If

        ' 3. Add the result (Default or OpenAI) to the Memory Cache for this run
        On Error Resume Next ' Handle potential error adding to dictionary
        dictCache(companyName) = finalRecap
        If Err.Number <> 0 Then
            LogEvt PROC_NAME, lgERROR, "Error adding '" & companyName & "' to memory cache: " & Err.Description ' Use lgERROR
            TraceEvt lvlERROR, PROC_NAME, "Error adding to memory cache", "Company=" & companyName & ", Err=" & Err.Description
            Err.Clear
        Else
             TraceEvt lvlDET, PROC_NAME, "Added to memory cache", "Company=" & companyName
        End If
        On Error GoTo 0 ' Restore default error handling
    End If

    GetCompanyRecap = finalRecap
End Function

Private Function GetCompanyRecapOpenAI(companyName As String) As String
    ' Purpose: Calls OpenAI API to get a company summary. Includes error handling & logging.
    Dim apiKey As String, result As String, http As Object, url As String, jsonPayload As String, jsonResponse As String
    Const PROC_NAME As String = "GetCompanyRecapOpenAI"
    GetCompanyRecapOpenAI = "" ' Default return value

    ' Double-check maintainer status (though already checked by caller)
    If Not IsMaintainerUser() Then
         LogEvt PROC_NAME, lgINFO, "Skipped OpenAI Call: Not Maintainer User.", "Company=" & companyName ' Use lgINFO
         TraceEvt lvlINFO, PROC_NAME, "Skipped: Not Maintainer" , "Company=" & companyName
        Exit Function ' Should not happen if called correctly, but safe
    End If

    ' Get API Key
    apiKey = GetAPIKey() ' Assumes GetAPIKey logs its own errors/warnings
    If apiKey = "" Then
        ' GetAPIKey function should have logged the reason
        GetCompanyRecapOpenAI = "Error: API Key Not Configured" ' Return error string
        TraceEvt lvlERROR, PROC_NAME, "Skipped: API Key Not Found/Configured", "Company=" & companyName
        Exit Function
    End If

    On Error GoTo OpenAIErrorHandler

    ' Prepare Request
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

    ' Send Request
    LogEvt PROC_NAME, lgDETAIL, "Sending request...", "Company=" & companyName & ", Model=" & modelName ' Use lgDETAIL
    TraceEvt lvlDET, PROC_NAME, "Sending request...", "Company=" & companyName & ", Model=" & modelName
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False ' Synchronous call
    http.setTimeouts OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send jsonPayload

    ' Process Response
    LogEvt PROC_NAME, lgDETAIL, "Response Received.", "Company=" & companyName & ", Status=" & http.Status ' Use lgDETAIL
    TraceEvt lvlDET, PROC_NAME, "Response Received", "Company=" & companyName & ", Status=" & http.Status

    If http.Status = 200 Then
        jsonResponse = http.responseText
        ' Attempt to parse the response JSON to extract content
        Const CONTENT_TAG As String = """content"":"""
        Dim contentStart As Long, contentEnd As Long, searchStart As Long
        searchStart = InStr(1, jsonResponse, """role"":""assistant""") ' Find assistant message block
        If searchStart > 0 Then
            contentStart = InStr(searchStart, jsonResponse, CONTENT_TAG)
            If contentStart > 0 Then
                contentStart = contentStart + Len(CONTENT_TAG)
                contentEnd = InStr(contentStart, jsonResponse, """") ' Find closing quote
                If contentEnd > contentStart Then
                    result = Mid$(jsonResponse, contentStart, contentEnd - contentStart)
                    result = JsonUnescape(result) ' Unescape special chars
                    LogEvt PROC_NAME, lgINFO, "OpenAI SUCCESS.", "Company=" & companyName ' Use lgINFO
                    TraceEvt lvlINFO, PROC_NAME, "OpenAI SUCCESS", "Company=" & companyName
                Else
                     result = "Error: Parse Fail (End Quote)"
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 500) ' Use lgERROR
                     TraceEvt lvlERROR, PROC_NAME, "Parse Fail (End Quote)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                End If
            Else
                ' Content tag not found, check if it's an error object from OpenAI
                If InStr(1, jsonResponse, """error""", vbTextCompare) > 0 Then
                     result = "Error: API returned error object."
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
                     TraceEvt lvlERROR, PROC_NAME, "API returned error object", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                Else
                     result = "Error: Parse Fail (Start Tag)"
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
                     TraceEvt lvlERROR, PROC_NAME, "Parse Fail (Start Tag)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                End If
            End If
        Else
             result = "Error: Parse Fail (No Assistant Role)"
             LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
             TraceEvt lvlERROR, PROC_NAME, "Parse Fail (No Assistant Role)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
        End If
    Else
        ' HTTP Error
        result = "Error: API Call Failed - Status " & http.Status & " - " & http.statusText
         LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(http.responseText, 500) ' Use lgERROR
         TraceEvt lvlERROR, PROC_NAME, "API Call Failed", "Company=" & companyName & ", Status=" & http.Status & ", ResponseStart=" & Left(http.responseText, 100)
    End If

    ' Cleanup and Finalize
    Set http = Nothing
    If Len(result) > RECAP_MAX_LEN Then result = Left$(result, RECAP_MAX_LEN - 3) & "..." ' Ensure truncation fits
    GetCompanyRecapOpenAI = Trim(result) ' Return the parsed/error string
    Exit Function

OpenAIErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     LogEvt PROC_NAME, lgERROR, "VBA Exception during OpenAI Call: " & errDesc, "Company=" & companyName ' Use lgERROR
     TraceEvt lvlERROR, PROC_NAME, "VBA Exception", "Company=" & companyName & ", Err=" & Err.Number & " - " & errDesc
    GetCompanyRecapOpenAI = "Error: VBA Exception - " & errDesc ' Return VBA error string
    If Not http Is Nothing Then Set http = Nothing ' Clean up object on error
End Function

Private Sub SaveCompanyCache(wsCache As Worksheet)
    ' Purpose: Saves the in-memory company cache back to the sheet.
    Dim key As Variant, i As Long, outputArr() As Variant, saveCount As Long
    Const PROC_NAME As String = "SaveCompanyCache"

    If dictCache Is Nothing Or dictCache.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "In-memory cache empty, skipping save." ' Use lgINFO
        TraceEvt lvlINFO, PROC_NAME, "Skipped saving empty cache"
        Exit Sub
    End If

    On Error GoTo CacheSaveError
    saveCount = dictCache.Count
    TraceEvt lvlINFO, PROC_NAME, "Saving cache to sheet", "Sheet=" & wsCache.Name & ", Items=" & saveCount
    ReDim outputArr(1 To saveCount, 1 To 3) ' CompanyName, RecapText, LastUpdated

    ' Populate the output array
    i = 1
    For Each key In dictCache.Keys
        outputArr(i, 1) = key
        outputArr(i, 2) = dictCache(key)
        outputArr(i, 3) = Now ' Timestamp the update
        i = i + 1
    Next key

    ' Prepare sheet and write data
    Dim previousEnableEvents As Boolean: previousEnableEvents = Application.EnableEvents
    Dim previousCalculation As XlCalculation: previousCalculation = Application.Calculation
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    With wsCache
        ' Clear existing data (excluding header)
        .Range("A2:C" & .Rows.Count).ClearContents
        If saveCount > 0 Then
            ' Write the new data
            .Range("A2").Resize(saveCount, 3).Value = outputArr
            ' Format the timestamp column
            .Range("C2").Resize(saveCount, 1).NumberFormat = "m/d/yyyy h:mm AM/PM"
            ' Autofit columns after writing
             On Error Resume Next ' Autofit might fail on hidden sheets sometimes
            .Columns("A:C").AutoFit
             On Error GoTo CacheSaveError ' Restore handler
        End If
    End With
    LogEvt PROC_NAME, lgINFO, "Saved " & saveCount & " items to cache sheet." ' Use lgINFO
    TraceEvt lvlINFO, PROC_NAME, "Cache save complete", "ItemsSaved=" & saveCount

CacheSaveExit: ' Label for normal exit and error exit cleanup
    ' Restore application settings
    Application.EnableEvents = previousEnableEvents
    Application.Calculation = previousCalculation
    Exit Sub

CacheSaveError:
     LogEvt PROC_NAME, lgERROR, "Error saving cache to sheet '" & wsCache.Name & "': " & Err.Description ' Use lgERROR
     TraceEvt lvlERROR, PROC_NAME, "Error saving cache", "Sheet=" & wsCache.Name & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "Error saving company cache to sheet '" & wsCache.Name & "': " & Err.Description, vbExclamation, "Cache Save Error"
    Resume CacheSaveExit ' Attempt to restore settings even after error
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
    Set loCol = tbl.ListColumns(colName) ' Assumes base name, not Name#Index needed here
    On Error GoTo SortErr ' Restore error handler
    If loCol Is Nothing Then
        LogEvt "Sort", lgWARN, "Column '" & colName & "' not found in table '" & tbl.Name & "' – sort skipped" ' Use lgWARN
        TraceEvt lvlWARN, "SortDataTable", "Column not found, sort skipped", "Table='" & tbl.Name & "', Column='" & colName & "'"
        Exit Sub
    End If
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=loCol.DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    LogEvt "Sort", lgINFO, "Table '" & tbl.Name & "' sorted by " & _
           colName & IIf(sortOrder = xlDescending, " (Desc)", " (Asc)") ' Use lgINFO
    TraceEvt lvlINFO, "SortDataTable", "Table sorted", "Table='" & tbl.Name & "', Column='" & colName & "', Order=" & IIf(sortOrder = xlDescending, "Desc", "Asc")
    Exit Sub
SortErr:
    LogEvt "Sort", lgERROR, "Error sorting table '" & tbl.Name & "' by column '" & colName & "': " & Err.Description ' Use lgERROR
    TraceEvt lvlERROR, "SortDataTable", "Error sorting table", "Table='" & tbl.Name & "', Column='" & colName & "', Err=" & Err.Number & " - " & Err.Description
End Sub

' --- CheckKeywords (Using RegExp) ---
Private Function CheckKeywords(textToCheck As String, keywordColl As Collection) As Boolean
    Dim kw As Variant
    CheckKeywords = False
    If keywordColl Is Nothing Or keywordColl.Count = 0 Or Len(Trim(textToCheck)) = 0 Then Exit Function

    ' --- Initialize RegExp object (Late Binding) ---
    If regex Is Nothing Then Set regex = CreateObject("VBScript.RegExp")

    ' --- Build pattern and test ---
    On Error GoTo CheckKeywordsErrorHandler ' Handle errors during RegExp or pattern building

    ' Build the pattern: (keyword1|keyword2|keyword3)
    ' Need to escape any special regex characters within keywords if they exist
    Dim patternBuilder As Object: Set patternBuilder = CreateObject("System.Collections.ArrayList") ' Use ArrayList for dynamic add
    For Each kw In keywordColl
        ' Basic escaping for common characters, might need more robust escaping if keywords are complex
        Dim escapedKw As String: escapedKw = CStr(kw)
        escapedKw = Replace(escapedKw, "\", "\\")
        escapedKw = Replace(escapedKw, ".", "\.")
        escapedKw = Replace(escapedKw, "|", "\|")
        escapedKw = Replace(escapedKw, "(", "\(")
        escapedKw = Replace(escapedKw, ")", "\)")
        escapedKw = Replace(escapedKw, "[", "\[")
        escapedKw = Replace(escapedKw, "]", "\]")
        escapedKw = Replace(escapedKw, "*", "\*")
        escapedKw = Replace(escapedKw, "+", "\+")
        escapedKw = Replace(escapedKw, "?", "\?")
        escapedKw = Replace(escapedKw, "{", "\{")
        escapedKw = Replace(escapedKw, "}", "\}")
        escapedKw = Replace(escapedKw, "^", "\^")
        escapedKw = Replace(escapedKw, "$", "\$")
        patternBuilder.Add escapedKw
    Next kw

    If patternBuilder.Count = 0 Then GoTo CheckKeywordsExit ' No valid keywords to build pattern

    regex.Pattern = Join(patternBuilder.ToArray(), "|") ' Join keywords with OR operator
    regex.IgnoreCase = True ' Case-insensitive match
    regex.Global = False    ' Only need to find one match

    ' Test the input string against the pattern
    CheckKeywords = regex.Test(textToCheck)

CheckKeywordsExit:
    Set patternBuilder = Nothing
    Exit Function

CheckKeywordsErrorHandler:
    LogEvt "CheckKeywords", lgERROR, "Error during RegExp keyword check: " & Err.Description ' Use lgERROR
    TraceEvt lvlERROR, "CheckKeywords", "RegExp Error", "Err=" & Err.Number & " - " & Err.Description
    Debug.Print Time & " - ERROR in CheckKeywords RegExp: " & Err.Description
    CheckKeywords = False ' Return False on error
    Resume CheckKeywordsExit ' Go to cleanup
End Function

' --- GetWeightFromDict (Helper) ---
Private Function GetWeightFromDict(dict As Object, key As String, defaultWeight As Double) As Double
    ' Purpose: Safely retrieves a weight (Double) from a dictionary, using default if key not found or value is invalid.
    If dict Is Nothing Then GetWeightFromDict = defaultWeight: Exit Function ' Handle Nothing dictionary object

    Dim value As Variant
    On Error Resume Next ' Suppress errors during dictionary access/conversion

    If dict.Exists(key) Then
        value = dict(key)
        If IsNumeric(value) Then
            GetWeightFromDict = CDbl(value) ' Convert valid numeric value to Double
            If Err.Number <> 0 Then GetWeightFromDict = defaultWeight: Err.Clear ' Use default if CDbl fails (overflow?)
        Else
            GetWeightFromDict = defaultWeight ' Value exists but is not numeric
        End If
    Else
        GetWeightFromDict = defaultWeight ' Key does not exist
    End If

    On Error GoTo 0 ' Restore default error handling
End Function

' --- WriteResultsToArray (Helper) ---
Private Sub WriteResultsToArray(ByRef dataArr As Variant, ByVal rowIdx As Long, ByVal cols As Object, ByVal scoreResult As Variant, ByVal recap As String)
    ' Purpose: Writes the calculated score results and recap into the correct columns of the data array for a specific row.
    Const PROC_NAME As String = "WriteResultsToArray"
    ' Use SafeGetColIndex to handle potential missing columns gracefully, though AddScoreColumnsIfNeeded should prevent this.
    Dim scoreCol As Long, catCol As Long, acCol As Long, pcCol As Long, kwCol As Long, stCol As Long
    Dim ptCol As Long, glCol As Long, nfCol As Long, synCol As Long, pctCol As Long, recapCol As Long

    scoreCol = SafeGetColIndex(cols, "Final_Score")
    catCol = SafeGetColIndex(cols, "Category")
    acCol = SafeGetColIndex(cols, "AC_Wt")
    pcCol = SafeGetColIndex(cols, "PC_Wt")
    kwCol = SafeGetColIndex(cols, "KW_Wt")
    stCol = SafeGetColIndex(cols, "ST_Wt")
    ptCol = SafeGetColIndex(cols, "PT_Wt")
    glCol = SafeGetColIndex(cols, "GL_Wt")
    nfCol = SafeGetColIndex(cols, "NF_Calc")
    synCol = SafeGetColIndex(cols, "Synergy_Calc")
    pctCol = SafeGetColIndex(cols, "Score_Percent")
    recapCol = SafeGetColIndex(cols, "CompanyRecap")

    ' Suppress errors during array write, log if any column index was invalid
    On Error Resume Next
    If scoreCol > 0 Then dataArr(rowIdx, scoreCol) = scoreResult(0)
    If catCol > 0 Then dataArr(rowIdx, catCol) = scoreResult(1)
    If acCol > 0 Then dataArr(rowIdx, acCol) = scoreResult(2)
    If pcCol > 0 Then dataArr(rowIdx, pcCol) = scoreResult(3)
    If kwCol > 0 Then dataArr(rowIdx, kwCol) = scoreResult(4)
    If stCol > 0 Then dataArr(rowIdx, stCol) = scoreResult(5)
    If ptCol > 0 Then dataArr(rowIdx, ptCol) = scoreResult(6)
    If glCol > 0 Then dataArr(rowIdx, glCol) = scoreResult(7)
    If nfCol > 0 Then dataArr(rowIdx, nfCol) = scoreResult(8)
    If synCol > 0 Then dataArr(rowIdx, synCol) = scoreResult(9)
    If pctCol > 0 Then dataArr(rowIdx, pctCol) = scoreResult(0) ' Store raw decimal for percentage format
    If recapCol > 0 Then dataArr(rowIdx, recapCol) = recap

    If Err.Number <> 0 Then
         LogEvt PROC_NAME, lgERROR, "Error writing results to array row " & rowIdx & ": " & Err.Description ' Use lgERROR
         TraceEvt lvlERROR, PROC_NAME, "Error writing to array", "Row=" & rowIdx & ", Err=" & Err.Number & " - " & Err.Description
         Debug.Print Time & " - Error writing results to array row " & rowIdx & ": " & Err.Description: Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Sub

' --- ReorganizeColumns (Using Copy/Delete, Handles Duplicates) ---
Private Sub ReorganizeColumns(tbl As ListObject)
    ' Purpose: Reorganizes columns to match desiredOrder using a robust
    '          insert/copy/delete method that avoids ListColumn.Position
    '          errors on QueryTable-backed ListObjects. Handles duplicate headers.
    Const PROC_NAME As String = "ReorganizeColumns-CopyDel"
    Dim desiredOrder As Variant, i As Long, baseColName As String ' Use Base name for desired order
    Dim srcColLc As ListColumn, tmpColLc As ListColumn ' Use ListColumn explicitly
    Dim desiredPos As Long, currentHeaderMap As Object ' Dictionary mapping KEY to INDEX
    Dim col As ListColumn ' For initial header check
    Dim srcColKey As String ' Key used to find source column (may include #Index)
    Dim srcColIndex As Long ' Index of the source column to move

    ' Define the desired order of the *first few* key columns (use BASE names)
    desiredOrder = Array( _
        "K_Number", "DecisionDate", "Applicant", "DeviceName", "Contact", _
        "Score_Percent", "Category", "CompanyRecap", "FDA_Link") ' Add other key cols as needed

    ' --- Initial Setup ---
    Dim previousScreenUpdating As Boolean: previousScreenUpdating = Application.ScreenUpdating
    Dim previousEnableEvents As Boolean: previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ReorgErr

    TraceEvt lvlINFO, PROC_NAME, "Starting column reorganization (Copy/Delete method)...", "Table='" & tbl.Name & "', Initial Cols=" & tbl.ListColumns.Count

    ' --- Build a map of current headers (Key -> Index), handling duplicates ---
    Set currentHeaderMap = CreateObject("Scripting.Dictionary")
    currentHeaderMap.CompareMode = vbTextCompare ' Case-insensitive for lookups
    Dim dupeCheckDict As Object: Set dupeCheckDict = CreateObject("Scripting.Dictionary")
    dupeCheckDict.CompareMode = vbBinaryCompare ' Case-sensitive check for exact duplicates

    For Each col In tbl.ListColumns
        Dim mapKey As String
        Dim colName As String: colName = Trim(col.Name)
        If Len(colName) = 0 Then
            TraceEvt lvlWARN, PROC_NAME, "Skipping blank header during map build", "Index=" & col.Index
            GoTo NextColMapBuild ' Skip blank headers
        End If

        If dupeCheckDict.Exists(colName) Then
            mapKey = colName & "#" & col.Index
            TraceEvt lvlDET, PROC_NAME, "Mapping duplicate header '" & colName & "' as '" & mapKey & "'", "Index=" & col.Index
        Else
            mapKey = colName
            dupeCheckDict.Add colName, 1
        End If

        If Not currentHeaderMap.Exists(mapKey) Then
             currentHeaderMap.Add mapKey, col.Index
        Else
             TraceEvt lvlWARN, PROC_NAME, "Duplicate key generated during map build: '" & mapKey & "'", "Index=" & col.Index
        End If
NextColMapBuild:
    Next col
    Set dupeCheckDict = Nothing
    TraceEvt lvlDET, PROC_NAME, "Built initial header map", "Count=" & currentHeaderMap.Count

    ' --- Loop through the desired columns IN ORDER ---
    For i = LBound(desiredOrder) To UBound(desiredOrder)
        baseColName = desiredOrder(i) ' The base name we want at this position
        desiredPos = i + 1 ' Target 1-based position
        srcColKey = "" ' Reset for each desired column
        srcColIndex = 0 ' Reset

        ' --- Find the *key* and *current index* of the column to move ---
        ' It should be the first instance matching baseColName whose current index >= desiredPos
        Dim currentItemKey As Variant
        Dim currentItemIndex As Long
        Dim foundKey As Boolean: foundKey = False

        For Each currentItemKey In currentHeaderMap.Keys
            currentItemIndex = currentHeaderMap(currentItemKey)
            Dim currentBaseName As String
            If InStr(currentItemKey, "#") > 0 Then
                currentBaseName = Left$(currentItemKey, InStr(currentItemKey, "#") - 1)
            Else
                currentBaseName = currentItemKey
            End If

            If StrComp(currentBaseName, baseColName, vbTextCompare) = 0 Then
                 ' Base name matches. Check if it's the one we should move.
                 If currentItemIndex >= desiredPos Then
                      ' Found a candidate. Select this one and stop searching for this base name.
                      srcColKey = currentItemKey
                      srcColIndex = currentItemIndex
                      foundKey = True
                      Exit For ' Move this specific key/index instance
                 End If
            End If
        Next currentItemKey

        ' --- Check if we found a column to move ---
        If Not foundKey Then
            TraceEvt lvlWARN, PROC_NAME, "Column base name '" & baseColName & "' not found or already placed. Skipping.", "DesiredPos=" & desiredPos
            GoTo NextIteration ' Skip to the next desired column
        End If

        TraceEvt lvlDET, PROC_NAME, "Identified source column to move", "Base='" & baseColName & "', Key='" & srcColKey & "', CurrentIndex=" & srcColIndex

        ' --- Get the source ListColumn object using its current INDEX ---
        Set srcColLc = Nothing ' Reset before trying to get
        On Error Resume Next ' Handle potential errors getting the column object by Index
        Set srcColLc = tbl.ListColumns(srcColIndex)
        On Error GoTo ReorgErr ' Restore error handler

        If srcColLc Is Nothing Then
             TraceEvt lvlERROR, PROC_NAME, "Failed to get ListColumn object for key '" & srcColKey & "' at index " & srcColIndex, "DesiredPos=" & desiredPos
             GoTo NextIteration ' Critical issue if index is invalid
        End If

        ' --- Check if it's already in the correct position ---
        If srcColIndex = desiredPos Then
            TraceEvt lvlDET, PROC_NAME, "Column '" & srcColKey & "' already in correct position.", "Index=" & desiredPos
            ' Remove from map so it's not considered again for moving
            currentHeaderMap.Remove srcColKey
            GoTo NextIteration ' Skip to the next desired column
        End If

        ' --- Perform the Insert/Copy/Delete ---
        TraceEvt lvlINFO, PROC_NAME, "Moving '" & srcColKey & "' (Base: '" & baseColName & "')", "From Index=" & srcColIndex & " To Position=" & desiredPos

        ' 1. Insert a new placeholder column at the target position
        Set tmpColLc = Nothing ' Reset
        On Error Resume Next ' Handle potential errors during Add
        Set tmpColLc = tbl.ListColumns.Add(Position:=desiredPos)
        Dim addErrNum As Long: addErrNum = Err.Number
        Dim addErrDesc As String: addErrDesc = Err.Description
        On Error GoTo ReorgErr ' Restore error handler

        If addErrNum <> 0 Or tmpColLc Is Nothing Then
            TraceEvt lvlERROR, PROC_NAME, "Failed to insert placeholder column at position " & desiredPos & " for '" & srcColKey & "'. Aborting move.", "SrcIndex=" & srcColIndex & ", Err=" & addErrNum & " - " & addErrDesc
            GoTo NextIteration ' Skip this column for now
        End If
        TraceEvt lvlSPAM, PROC_NAME, "Inserted placeholder", "Name='" & tmpColLc.Name & "', Pos=" & desiredPos & ", For='" & srcColKey & "'"

        ' 2. Copy header and data
        On Error Resume Next ' Handle potential errors during copy
        tmpColLc.Name = baseColName ' <<< SET THE HEADER TO THE DESIRED BASE NAME >>>
        If Err.Number <> 0 Then TraceEvt lvlWARN, PROC_NAME, "Error setting tmpCol name to '" & baseColName & "'", "Err=" & Err.Number & " - " & Err.Description: Err.Clear

        If srcColLc.DataBodyRange Is Nothing Then
             TraceEvt lvlWARN, PROC_NAME, "Source column '" & srcColKey & "' has no DataBodyRange to copy.", "SrcIndex=" & srcColIndex
        ElseIf tmpColLc.DataBodyRange Is Nothing Then
             TraceEvt lvlWARN, PROC_NAME, "Temp column '" & baseColName & "' has no DataBodyRange to paste into.", "TmpIndex=" & tmpColLc.Index
        Else
            tmpColLc.DataBodyRange.Value = srcColLc.DataBodyRange.Value ' Copy values
            If Err.Number <> 0 Then TraceEvt lvlWARN, PROC_NAME, "Error copying DataBodyRange for '" & srcColKey & "' to '" & baseColName & "'", "Err=" & Err.Number & " - " & Err.Description: Err.Clear Else TraceEvt lvlSPAM, PROC_NAME, "Copied data", "FromKey='" & srcColKey & "' ToName='" & tmpColLc.Name & "'"
        End If
        On Error GoTo ReorgErr ' Restore error handler

        ' 3. Delete the original source column
        Dim originalIndexBeforeDelete As Long: originalIndexBeforeDelete = srcColLc.Index ' Capture index just before delete
        On Error Resume Next ' Handle potential errors during delete
        srcColLc.Delete
        Dim delErrNum As Long: delErrNum = Err.Number
        Dim delErrDesc As String: delErrDesc = Err.Description
        On Error GoTo ReorgErr ' Restore error handler

        If delErrNum <> 0 Then
            TraceEvt lvlERROR, PROC_NAME, "Failed to delete original column (Key: '" & srcColKey & "') at index " & originalIndexBeforeDelete, "Err=" & delErrNum & " - " & delErrDesc
             ' This is problematic, might leave duplicates. Consider stopping.
        Else
            TraceEvt lvlSPAM, PROC_NAME, "Deleted original column", "Key='" & srcColKey & "', OriginalIndex=" & originalIndexBeforeDelete
        End If

        ' --- Update the header map after successful move ---
        ' Remove old entry for the moved key
        currentHeaderMap.Remove srcColKey
        ' Add the new column using its BASE name and its NEW index
        Dim newKey As String: newKey = baseColName
        ' Check if base name now conflicts with another column already in the map
        If currentHeaderMap.Exists(newKey) Then
            newKey = baseColName & "#" & tmpColLc.Index ' Use Name#Index if base name is now duplicated
            TraceEvt lvlDET, PROC_NAME, "Post-move map update: using key '" & newKey & "' due to base name conflict", "Index=" & tmpColLc.Index
        End If
        currentHeaderMap.Add newKey, tmpColLc.Index ' tmpColLc is now the actual column at the correct index

NextIteration:
        Set srcColLc = Nothing ' Clean up for next loop
        Set tmpColLc = Nothing
    Next i

    TraceEvt lvlINFO, PROC_NAME, "Column reorganization complete (Copy/Delete method).", "Final Cols=" & tbl.ListColumns.Count

TidyUp:
    ' Restore UI settings
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    ' Release objects
    Set srcColLc = Nothing
    Set tmpColLc = Nothing
    Set currentHeaderMap = Nothing
    Set col = Nothing
    Exit Sub

ReorgErr:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errSource As String: errSource = Err.Source
    Dim errMsg As String
    errMsg = "Err #" & errNum & " in " & errSource & ": " & errDesc
    Dim detailMsg As String
    detailMsg = "Current Action: Moving Base='" & baseColName & "' (Key='" & srcColKey & "') | Loop Index: " & i & _
                " | DesiredPos: " & desiredPos & _
                " | SrcCol Index: " & IIf(srcColLc Is Nothing, srcColIndex, srcColLc.Index) & _
                " | TmpCol Index: " & IIf(tmpColLc Is Nothing, "N/A", tmpColLc.Index) & _
                " | Table: '" & tbl.Name & "'"

    TraceEvt lvlERROR, PROC_NAME, errMsg, detailMsg
    LogEvt PROC_NAME, lgERROR, errMsg, detailMsg ' Use lgERROR
    Debug.Print Now & " - FATAL ERROR in " & PROC_NAME & ": " & errMsg & " | Details: " & detailMsg

    MsgBox "A critical error occurred during column reorganization:" & vbCrLf & vbCrLf & _
           errMsg & vbCrLf & detailMsg & vbCrLf & vbCrLf & _
           "The table layout might be incorrect.", vbCritical, "Reorganization Error"

    Resume TidyUp ' Attempt graceful exit and UI restore
End Sub

' --- FormatTableLook ---
Private Sub FormatTableLook(ws As Worksheet)
    ' Purpose: Applies consistent table style, alignment, widths, borders, and specific header formatting.
    Const PROC_NAME As String = "FormatTableLook"
    On Error GoTo FormatLookErrorHandler
    Dim tbl As ListObject
    Dim listCol As ListColumn
    Dim centerCols As Variant
    Dim wideCols As Variant
    Dim colName As Variant

    If ws.ListObjects.Count = 0 Then
        LogEvt "Formatting", lgWARN, "No table found on sheet '" & ws.Name & "' for FormatTableLook." ' Use lgWARN
        Exit Sub
    End If
    Set tbl = ws.ListObjects(1)
    TraceEvt lvlDET, PROC_NAME, "Applying format", "Sheet='" & ws.Name & "', Table='" & tbl.Name & "'"

    centerCols = Array("ProcTimeDays", "AC", "PC", "Category")
    wideCols = Array("DeviceName", "Applicant", "CompanyRecap") ' Not currently used, but defined

    LogEvt "Formatting", lgDETAIL, "Applying table style, alignment, widths, borders for table '" & tbl.Name & "'..." ' Use lgDETAIL

    ' Explicit Header Formatting
    TraceEvt lvlSPAM, PROC_NAME, "Applying header format"
    With tbl.HeaderRowRange
        .Interior.Color = RGB(31, 78, 121) ' Dark Blue
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Center specific data columns
    TraceEvt lvlSPAM, PROC_NAME, "Centering specific columns"
    For Each colName In centerCols
        On Error Resume Next ' Ignore if column doesn't exist
        Set listCol = Nothing
        Set listCol = tbl.ListColumns(colName) ' Assumes base name ok here
        If Not listCol Is Nothing Then
            listCol.DataBodyRange.HorizontalAlignment = xlCenter
            TraceEvt lvlSPAM, PROC_NAME, "Centered column", colName
        Else
            TraceEvt lvlWARN, PROC_NAME, "Column not found for centering", colName
            LogEvt "Formatting", lgDETAIL, "Column not found for centering: " & colName ' Use lgDETAIL
        End If
        On Error GoTo FormatLookErrorHandler ' Restore error handler
    Next colName
    Set listCol = Nothing

    ' Autofit all columns
    TraceEvt lvlSPAM, PROC_NAME, "Autofitting columns"
    On Error Resume Next
    tbl.Range.Columns.AutoFit
    If Err.Number <> 0 Then TraceEvt lvlWARN, PROC_NAME, "Error autofitting columns", Err.Description: Err.Clear
    On Error GoTo FormatLookErrorHandler

    ' Set specific widths
    TraceEvt lvlSPAM, PROC_NAME, "Setting specific column widths"
    On Error Resume Next
    If ColumnExistsInMap(GetColumnIndices(tbl.HeaderRowRange), "DeviceName") Then tbl.ListColumns("DeviceName").Range.ColumnWidth = 45
    If ColumnExistsInMap(GetColumnIndices(tbl.HeaderRowRange), "Applicant") Then tbl.ListColumns("Applicant").Range.ColumnWidth = 30
    If ColumnExistsInMap(GetColumnIndices(tbl.HeaderRowRange), "CompanyRecap") Then tbl.ListColumns("CompanyRecap").Range.ColumnWidth = 50
    If Err.Number <> 0 Then TraceEvt lvlWARN, PROC_NAME, "Error setting specific widths", Err.Description: Err.Clear
    On Error GoTo FormatLookErrorHandler

    ' Apply thin BLACK borders
    TraceEvt lvlSPAM, PROC_NAME, "Applying black borders"
    With tbl.Range.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    On Error Resume Next ' Apply to inside borders as well
    With tbl.Range.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = vbBlack
    End With
    With tbl.Range.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = vbBlack
    End With
    If Err.Number <> 0 Then TraceEvt lvlWARN, PROC_NAME, "Error applying borders", Err.Description: Err.Clear
    On Error GoTo FormatLookErrorHandler

    TraceEvt lvlINFO, PROC_NAME, "Completed OK", "Sheet='" & ws.Name & "'"
    LogEvt "Formatting", lgDETAIL, "Applied FormatTableLook with custom header and black borders." ' Use lgDETAIL
    Exit Sub

FormatLookErrorHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    TraceEvt lvlERROR, PROC_NAME, "Error applying table look formatting", "Sheet='" & ws.Name & "', Err=" & errNum & " - " & errDesc
    LogEvt "Formatting", lgERROR, "Error applying table look formatting on sheet '" & ws.Name & "': " & errDesc ' Use lgERROR
    Debug.Print Time & " - Error applying table look formatting on sheet '" & ws.Name & "': " & errDesc
    MsgBox "Error applying table formatting: " & errDesc, vbExclamation
End Sub

' --- FormatCategoryColors ---
Private Sub FormatCategoryColors(tblData As ListObject)
    Dim rng As Range
    On Error Resume Next
    Set rng = tblData.ListColumns("Category").DataBodyRange ' Assumes base name ok
    If rng Is Nothing Then Exit Sub
    On Error GoTo CatColorError
    LogEvt "Formatting", lgDETAIL, "Applying category conditional formatting..." ' Use lgDETAIL
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
    LogEvt "Formatting", lgDETAIL, "Applied category colors." ' Use lgDETAIL
    Exit Sub
CatColorError:
    LogEvt "Formatting", lgERROR, "Error applying category colors: " & Err.Description ' Use lgERROR
    TraceEvt lvlERROR, "FormatCategoryColors", "Error applying category colors", "Err=" & Err.Number & " - " & Err.Description
    Debug.Print Time & " - Error applying category colors: " & Err.Description
End Sub

' --- ApplyCondColor (Helper) ---
Private Sub ApplyCondColor(rng As Range, categoryText As String, fillColor As Long)
    With rng.FormatConditions.Add(Type:=xlTextString, String:=categoryText, TextOperator:=xlEqual)
        .Interior.Color = fillColor
        .Font.Color = IIf(GetBrightness(fillColor) < 130, vbWhite, vbBlack) ' Adjust font color for readability
        .StopIfTrue = False ' Allow other conditions if needed
    End With
End Sub

' --- GetBrightness (Helper) ---
Private Function GetBrightness(clr As Long) As Double
    ' Calculates perceived brightness of an RGB color (0-255 scale).
    On Error Resume Next
    GetBrightness = ((clr Mod 256) * 0.299) + (((clr \ 256) Mod 256) * 0.587) + (((clr \ 65536) Mod 256) * 0.114)
    If Err.Number <> 0 Then GetBrightness = 128 ' Default brightness on error
    On Error GoTo 0
End Function

' --- ApplyNumberFormats ---
Private Sub ApplyNumberFormats(tblData As ListObject)
    On Error Resume Next ' Ignore errors if a column doesn't exist
    Const PROC_NAME As String = "ApplyNumberFormats"
    LogEvt PROC_NAME, lgDETAIL, "Applying number formats..." ' Use lgDETAIL
    TraceEvt lvlDET, PROC_NAME, "Applying number formats", "Table=" & tblData.Name
    If ColumnExistsInMap(GetColumnIndices(tblData.HeaderRowRange), "Score_Percent") Then tblData.ListColumns("Score_Percent").DataBodyRange.NumberFormat = "0.0%"
    If ColumnExistsInMap(GetColumnIndices(tblData.HeaderRowRange), "Final_Score") Then tblData.ListColumns("Final_Score").DataBodyRange.NumberFormat = "0.000"
    Dim scoreWtCols As Variant: scoreWtCols = Array("AC_Wt", "PC_Wt", "KW_Wt", "ST_Wt", "PT_Wt", "GL_Wt", "NF_Calc", "Synergy_Calc")
    Dim colName As Variant
    For Each colName In scoreWtCols
        If ColumnExistsInMap(GetColumnIndices(tblData.HeaderRowRange), colName) Then tblData.ListColumns(colName).DataBodyRange.NumberFormat = "0.00"
    Next colName
    If ColumnExistsInMap(GetColumnIndices(tblData.HeaderRowRange), "ProcTimeDays") Then tblData.ListColumns("ProcTimeDays").DataBodyRange.NumberFormat = "0"
    LogEvt PROC_NAME, lgDETAIL, "Number formats applied." ' Use lgDETAIL
    TraceEvt lvlDET, PROC_NAME, "Number formats applied"
    On Error GoTo 0
End Sub

' --- CreateShortNamesAndComments ---
Private Sub CreateShortNamesAndComments(tblData As ListObject)
    Dim devNameCol As ListColumn, devNameRange As Range, cell As Range
    Dim originalName As String, shortName As String
    Const PROC_NAME As String = "CreateShortNamesAndComments"
    On Error GoTo ShortNameErrorHandler

    On Error Resume Next ' Check if DeviceName column exists
    Set devNameCol = tblData.ListColumns("DeviceName") ' Assumes base name ok
    If devNameCol Is Nothing Then
        LogEvt PROC_NAME, lgWARN, "DeviceName column not found, skipping short names/comments." ' Use lgWARN
        TraceEvt lvlWARN, PROC_NAME, "DeviceName column not found, skipping."
        Exit Sub
    End If
    Set devNameRange = devNameCol.DataBodyRange
    If devNameRange Is Nothing Then Exit Sub ' Exit if no data rows
    On Error GoTo ShortNameErrorHandler ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Applying smart short names and comments..." ' Use lgDETAIL
    TraceEvt lvlDET, PROC_NAME, "Applying smart short names/comments", "Column=DeviceName, Rows=" & devNameRange.Rows.Count

    Dim previousScreenUpdating As Boolean: previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Optimize updates
    Dim cellCount As Long: cellCount = 0

    For Each cell In devNameRange.Cells
        cellCount = cellCount + 1
        On Error Resume Next ' Handle errors per cell (e.g., locked cell)
        If Not cell.Comment Is Nothing Then cell.Comment.Delete ' Clear existing comment first
        originalName = Trim(CStr(cell.Value))
        If Len(originalName) > SHORT_NAME_MAX_LEN Then
            ' Simple shortening logic (can be refined)
            If InStr(1, originalName, "(") > 10 Then ' Shorten before first parenthesis if reasonable length
                shortName = Trim(Left(originalName, InStr(1, originalName, "(") - 1))
            ElseIf InStr(1, originalName, ";") > 10 Then ' Shorten before first semicolon
                shortName = Trim(Left(originalName, InStr(1, originalName, ";") - 1)) & SHORT_NAME_ELLIPSIS
            ElseIf InStr(1, originalName, ",") > 25 Then ' Shorten before first comma if > 25 chars
                shortName = Trim(Left(originalName, InStr(1, originalName, ",") - 1)) & SHORT_NAME_ELLIPSIS
            Else ' Default hard truncation
                shortName = Left$(originalName, SHORT_NAME_MAX_LEN - Len(SHORT_NAME_ELLIPSIS)) & SHORT_NAME_ELLIPSIS
            End If
            ' Ensure minimum length and avoid tiny fragments
            If Len(shortName) < 10 Then
                shortName = Left$(originalName, SHORT_NAME_MAX_LEN - Len(SHORT_NAME_ELLIPSIS)) & SHORT_NAME_ELLIPSIS
            End If

            ' Update cell value only if it changed
            If cell.Value <> shortName Then
                cell.Value = shortName
                TraceEvt lvlSPAM, PROC_NAME, "Shortened device name", "Row=" & cell.Row & ", OrigLen=" & Len(originalName) & ", NewLen=" & Len(shortName)
            End If

            ' Add the original name as a comment
            cell.AddComment Text:=originalName
            If Err.Number = 0 Then
                 On Error Resume Next ' Handle potential error setting AutoSize
                cell.Comment.Shape.TextFrame.AutoSize = True
                 On Error GoTo ShortNameErrorHandler ' Restore handler
            Else
                LogEvt PROC_NAME, lgWARN, "Could not add comment to " & cell.Address(External:=True) & ": " & Err.Description ' Use lgWARN
                TraceEvt lvlWARN, PROC_NAME, "Could not add comment", "Cell=" & cell.Address & ", Err=" & Err.Number & " - " & Err.Description
                Err.Clear ' Clear error for this cell
            End If
        End If
        On Error GoTo ShortNameErrorHandler ' Restore main handler for loop
    Next cell

    Application.ScreenUpdating = previousScreenUpdating ' Restore setting
    LogEvt PROC_NAME, lgINFO, "Smart short names/comments processing complete.", "Processed=" & cellCount ' Use lgINFO
    TraceEvt lvlINFO, PROC_NAME, "Short names/comments complete", "Processed=" & cellCount
    Exit Sub

ShortNameErrorHandler:
    Application.ScreenUpdating = True ' Ensure screen updating is re-enabled on error exit
    LogEvt PROC_NAME, lgERROR, "Error applying smart short names/comments: " & Err.Description ' Use lgERROR
    TraceEvt lvlERROR, PROC_NAME, "Error applying short names/comments", "Err=" & Err.Number & " - " & Err.Description
    MsgBox "Error applying smart device names/comments: " & Err.Description, vbExclamation, "Short Name Error"
End Sub

' --- FreezeHeaderAndKeyCols (Still Disabled) ---
Private Sub FreezeHeaderAndKeyCols(ws As Worksheet)
    Const PROC_NAME As String = "FreezeHeaderAndKeyCols"
    TraceEvt lvlINFO, PROC_NAME, "Freeze panes currently disabled.", "Sheet='" & ws.Name & "'"
    Exit Sub ' <<< DISABLE FREEZING ENTIRELY (per feedback) >>>
    ' ... (rest of the disabled code remains unchanged) ...
End Sub

' --- ArchiveMonth ---
Private Sub ArchiveMonth(wsDataSource As Worksheet, archiveSheetName As String)
    ' Purpose: Creates an archive copy of the data source sheet, converting to values.
    Dim wsArchive As Worksheet
    Const PROC_NAME As String = "ArchiveMonth"
    Dim previousDisplayAlerts As Boolean: previousDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ArchiveErrorHandler
    Application.DisplayAlerts = False ' Prevent prompts (e.g., delete sheet)
    LogEvt PROC_NAME, lgINFO, "Starting archive process for: " & archiveSheetName ' Use lgINFO
    TraceEvt lvlINFO, PROC_NAME, "Starting archive", "Source=" & wsDataSource.Name & ", Target=" & archiveSheetName

    ' Copy the sheet
    wsDataSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsArchive = ActiveSheet ' The newly copied sheet

    ' Rename the copied sheet
    On Error Resume Next ' Handle potential naming conflicts or invalid characters
    wsArchive.Name = Left(archiveSheetName, 31) ' Ensure name <= 31 chars
    If Err.Number <> 0 Then
        Dim fallbackName As String: fallbackName = "Archive_Error_" & Format(Now(), "yyyyMMdd_HHmmss")
        LogEvt PROC_NAME, lgWARN, "Rename to '" & archiveSheetName & "' failed (Err#" & Err.Number & "). Using fallback: " & fallbackName ' Use lgWARN
        TraceEvt lvlWARN, PROC_NAME, "Rename failed, using fallback", "Target=" & archiveSheetName & ", Fallback=" & fallbackName & ", Err=" & Err.Number
        Err.Clear
        wsArchive.Name = fallbackName
    End If
    On Error GoTo ArchiveErrorHandler ' Restore main handler

    ' Convert formulas to values
    If wsArchive.UsedRange.Cells.CountLarge > 1 Then ' Check if there's anything to convert
         On Error Resume Next ' Handle error if range is protected etc.
        wsArchive.UsedRange.Value = wsArchive.UsedRange.Value
        If Err.Number = 0 Then
             LogEvt PROC_NAME, lgDETAIL, "Converted formulas to values on sheet: " & wsArchive.Name ' Use lgDETAIL
             TraceEvt lvlDET, PROC_NAME, "Converted formulas to values", "Sheet=" & wsArchive.Name
        Else
             LogEvt PROC_NAME, lgWARN, "Could not convert formulas to values on sheet: " & wsArchive.Name & ", Err=" & Err.Description ' Use lgWARN
             TraceEvt lvlWARN, PROC_NAME, "Failed converting formulas to values", "Sheet=" & wsArchive.Name & ", Err=" & Err.Number
             Err.Clear
        End If
         On Error GoTo ArchiveErrorHandler ' Restore handler
    End If

    ' Unlist the table
    If wsArchive.ListObjects.Count > 0 Then
        On Error Resume Next ' Handle error unlisting
        wsArchive.ListObjects(1).Unlist
        If Err.Number = 0 Then
             LogEvt PROC_NAME, lgDETAIL, "Unlisted table on archive sheet: " & wsArchive.Name ' Use lgDETAIL
             TraceEvt lvlDET, PROC_NAME, "Unlisted table", "Sheet=" & wsArchive.Name
        Else
             LogEvt PROC_NAME, lgWARN, "Could not unlist table on sheet: " & wsArchive.Name & ", Err=" & Err.Description ' Use lgWARN
             TraceEvt lvlWARN, PROC_NAME, "Failed unlisting table", "Sheet=" & wsArchive.Name & ", Err=" & Err.Number
             Err.Clear
        End If
        On Error GoTo ArchiveErrorHandler ' Restore handler
    End If

    ' Optional: Clear Comments
    ' On Error Resume Next: wsArchive.Cells.ClearComments: LogEvt PROC_NAME, lgDETAIL, "Cleared comments from archive." : TraceEvt lvlDET, PROC_NAME, "Cleared comments": On Error GoTo ArchiveErrorHandler

    ' Optional: Protect Sheet
    ' On Error Resume Next
    ' wsArchive.Protect Password:="YourPassword", UserInterfaceOnly:=True ' Add other protection options as needed
    ' If Err.Number = 0 Then LogEvt PROC_NAME, lgDETAIL, "Protected archive sheet: " & wsArchive.Name: TraceEvt lvlDET, PROC_NAME, "Protected sheet" Else LogEvt PROC_NAME, lgWARN, "Failed to protect sheet": TraceEvt lvlWARN, PROC_NAME, "Protection failed"
    ' Err.Clear: On Error GoTo ArchiveErrorHandler

    LogEvt PROC_NAME, lgINFO, "Successfully archived data to sheet: " & wsArchive.Name ' Use lgINFO
    TraceEvt lvlINFO, PROC_NAME, "Archive successful", "Sheet=" & wsArchive.Name
    Application.DisplayAlerts = previousDisplayAlerts ' Restore alerts
    Exit Sub

ArchiveErrorHandler:
    Dim errDesc As String: errDesc = Err.Description: Dim errNum As Long: errNum = Err.Number
    Application.DisplayAlerts = True ' Ensure alerts are back on after error
    LogEvt PROC_NAME, lgERROR, "Error during archiving for '" & archiveSheetName & "': " & errDesc & " (#" & errNum & ")" ' Use lgERROR
    TraceEvt lvlERROR, PROC_NAME, "Error during archive", "Target=" & archiveSheetName & ", Err=" & errNum & " - " & errDesc
    MsgBox "Error during archiving process for sheet '" & archiveSheetName & "': " & vbCrLf & errDesc, vbCritical, "Archive Error"
    ' Attempt to delete the partially created/failed archive sheet
    If Not wsArchive Is Nothing Then
        If wsArchive.Name <> wsDataSource.Name Then ' Check if it's the copied sheet
            On Error Resume Next ' Ignore errors during cleanup delete
            wsArchive.Delete
            If Err.Number = 0 Then TraceEvt lvlWARN, PROC_NAME, "Deleted partial archive sheet due to error."
            On Error GoTo 0 ' Restore normal error handling after attempted delete
        End If
    End If
End Sub

' --- GetAPIKey ---
Private Function GetAPIKey() As String
    ' Purpose: Reads the OpenAI API key from a specified file path.
    Dim fso As Object, ts As Object, keyPath As String, WshShell As Object, fileContent As String: fileContent = ""
    Const PROC_NAME As String = "GetAPIKey"
    On Error GoTo KeyError

    ' Expand environment variables in the path
    Set WshShell = CreateObject("WScript.Shell")
    keyPath = WshShell.ExpandEnvironmentStrings(API_KEY_FILE_PATH)
    Set WshShell = Nothing
    TraceEvt lvlDET, PROC_NAME, "Resolved API Key Path", keyPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(keyPath) Then
        Set ts = fso.OpenTextFile(keyPath, 1) ' ForReading
        If Not ts.AtEndOfStream Then fileContent = ts.ReadAll
        ts.Close
        If Len(Trim(fileContent)) > 0 Then
             LogEvt PROC_NAME, lgDETAIL, "API Key read successfully." ' Use lgDETAIL
             TraceEvt lvlDET, PROC_NAME, "API Key read successfully"
        Else
             LogEvt PROC_NAME, lgWARN, "API Key file exists but is empty.", "Path=" & keyPath ' Use lgWARN
             TraceEvt lvlWARN, PROC_NAME, "API Key file empty", "Path=" & keyPath
        End If
    Else
         LogEvt PROC_NAME, lgWARN, "API Key file not found.", "Path=" & keyPath ' Use lgWARN
         TraceEvt lvlWARN, PROC_NAME, "API Key file not found", "Path=" & keyPath
        Debug.Print Time & " - WARNING: API Key file not found at specified path: " & keyPath
    End If
    GoTo KeyExit

KeyError:
     LogEvt PROC_NAME, lgERROR, "Error reading API Key from '" & keyPath & "': " & Err.Description ' Use lgERROR
     TraceEvt lvlERROR, PROC_NAME, "Error reading API Key file", "Path=" & keyPath & ", Err=" & Err.Number & " - " & Err.Description
    Debug.Print Time & " - ERROR reading API Key from '" & keyPath & "': " & Err.Description

KeyExit:
    GetAPIKey = Trim(fileContent)
    ' Clean up objects
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    On Error GoTo 0 ' Restore default error handling
End Function

' --- JsonEscape ---
Private Function JsonEscape(strInput As String) As String
    ' Purpose: Escapes characters in a string for safe inclusion in a JSON payload.
    strInput = Replace(strInput, "\", "\\")   ' Escape backslashes FIRST
    strInput = Replace(strInput, """", "\""") ' Escape double quotes
    strInput = Replace(strInput, vbCrLf, "\n") ' Replace CRLF with \n
    strInput = Replace(strInput, vbCr, "\n")   ' Replace CR with \n
    strInput = Replace(strInput, vbLf, "\n")   ' Replace LF with \n
    strInput = Replace(strInput, vbTab, "\t")  ' Replace Tab with \t
    ' Add other escapes if needed (e.g., for control characters < U+0020)
    JsonEscape = strInput
End Function

' --- JsonUnescape ---
Private Function JsonUnescape(strInput As String) As String
    ' Purpose: Unescapes characters in a string retrieved from a JSON payload.
    strInput = Replace(strInput, "\n", vbCrLf) ' Convert \n back to CRLF
    strInput = Replace(strInput, "\t", vbTab)  ' Convert \t back to Tab
    strInput = Replace(strInput, "\""", """") ' Unescape double quotes
    strInput = Replace(strInput, "\\", "\")   ' Unescape backslashes LAST
    JsonUnescape = strInput
End Function

' --- EnsureUIOn ---
Private Sub EnsureUIOn()
    ' Purpose: Guarantees that key Application settings affecting UI are restored.
    '          Call this from error handlers and clean exit points.
    Const PROC_NAME As String = "EnsureUIOn"
    On Error Resume Next ' Don't let this routine itself cause cascading errors
    Dim settingsChanged As Boolean: settingsChanged = False

    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True: settingsChanged = True
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic: settingsChanged = True
    If Not Application.EnableEvents Then Application.EnableEvents = True: settingsChanged = True
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True: settingsChanged = True
    If Application.StatusBar <> False Then Application.StatusBar = False: settingsChanged = True
    If Application.Cursor <> xlDefault Then Application.Cursor = xlDefault: settingsChanged = True

    If settingsChanged Then
        Debug.Print Time & " - EnsureUIOn: Application UI Settings Restored."
        TraceEvt lvlDET, PROC_NAME, "Application UI Settings Restored"
    End If
    On Error GoTo 0
End Sub

' ==========================================================================
' ===                        END OF MODULE                               ===
' ==========================================================================
