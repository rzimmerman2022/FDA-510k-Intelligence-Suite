' ==========================================================================
' Module      : mod_Cache
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module handles the caching mechanism for company summary
'               information ("recaps"). It maintains an in-memory dictionary
'               (dictCache) populated from a persistent worksheet cache
'               (defined by CACHE_SHEET_NAME in mod_Config). When a recap is
'               requested via GetCompanyRecap, it first checks the memory cache.
'               If not found, and if the user is a designated maintainer
'               (checked via mod_Utils.IsMaintainerUser) and OpenAI integration
'               is implicitly enabled, it attempts to fetch a summary from the
'               OpenAI API using GetCompanyRecapOpenAI. The fetched or default
'               recap is then added to the memory cache. The entire memory
'               cache can be saved back to the worksheet using SaveCompanyCache.
'               Helper functions manage API key retrieval and JSON string handling.
'
' Key Functions:
'               - LoadCompanyCache: Loads data from the cache sheet into the
'                 in-memory dictionary (dictCache).
'               - GetCompanyRecap: Retrieves a company recap, utilizing the
'                 cache and optionally calling OpenAI.
'               - SaveCompanyCache: Writes the current in-memory cache back
'                 to the persistent cache sheet.
'
' Private Helpers:
'               - GetCompanyRecapOpenAI: Handles the specific logic for calling
'                 the OpenAI API, including request formatting and response parsing.
'               - GetAPIKey: Reads the OpenAI API key from a configured file path.
'               - JsonEscape/JsonUnescape: Utilities for handling special characters
'                 in JSON strings.
'
' Dependencies: - mod_Logger: For logging cache operations, hits/misses, and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - mod_Config: For constants (cache sheet name, API URL, model,
'                 API key path, default recap text, timeouts, max lengths).
'               - mod_Utils: For checking maintainer status (IsMaintainerUser).
'               - Requires Scripting.Dictionary object.
'               - Requires MSXML2.ServerXMLHTTP.6.0 object for OpenAI calls.
'               - Requires Scripting.FileSystemObject for API key file reading.
'               - Requires WScript.Shell object for environment variable expansion.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit

' Module-level variable for the in-memory cache
Private dictCache As Object ' Scripting.Dictionary: Key=CompanyName, Value=RecapText

' --- Public Cache Management Functions ---

Public Sub LoadCompanyCache(wsCache As Worksheet)
    ' Purpose: Loads the persistent company cache from the sheet into memory.
    Dim lastRow As Long, i As Long, cacheData As Variant, loadedCount As Long
    Const PROC_NAME As String = "mod_Cache.LoadCompanyCache" ' Updated PROC_NAME
    Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Loading cache from sheet", "Sheet=" & wsCache.Name

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
     mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Cache loading complete", "ItemsLoaded=" & loadedCount
    On Error GoTo 0 ' Ensure normal error handling restored
    Exit Sub

CacheLoadError:
     LogEvt PROC_NAME, lgERROR, "Error reading cache data from sheet: " & Err.Description
     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error reading cache data", "Sheet=" & wsCache.Name & ", Err=" & Err.Number & " - " & Err.Description
     Err.Clear ' Clear error before resuming
     Resume ExitLoadCache ' Go to cleanup/logging part
End Sub

Public Function GetCompanyRecap(companyName As String, useOpenAI As Boolean) As String
    ' Purpose: Retrieves company recap, using memory cache, sheet cache, or optionally OpenAI.
    Dim finalRecap As String
    Const PROC_NAME As String = "mod_Cache.GetCompanyRecap" ' Updated PROC_NAME

    ' Initialize cache dictionary if it's not already set
    If dictCache Is Nothing Then Set dictCache = CreateObject("Scripting.Dictionary"): dictCache.CompareMode = vbTextCompare

    ' Handle invalid input
    If Len(Trim(companyName)) = 0 Then
        LogEvt PROC_NAME, lgWARN, "Invalid (empty) company name passed.", "RowContext=Unknown" ' Use lgWARN
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Invalid company name passed", "Name=''"
        GetCompanyRecap = "Invalid Applicant Name"
        Exit Function
    End If

    ' 1. Check Memory Cache (Fastest)
    If dictCache.Exists(companyName) Then
        finalRecap = dictCache(companyName)
        LogEvt PROC_NAME, lgDETAIL, "Memory Cache HIT.", "Company=" & companyName ' Use lgDETAIL
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Memory Cache Hit", "Company=" & companyName
    Else
        ' 2. Memory Cache MISS - Try OpenAI (if enabled) or use Default
        LogEvt PROC_NAME, lgDETAIL, "Memory Cache MISS.", "Company=" & companyName ' Use lgDETAIL
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Memory Cache Miss", "Company=" & companyName

        finalRecap = DEFAULT_RECAP_TEXT ' Assume default unless OpenAI succeeds (Requires mod_Config)

        If useOpenAI Then
            Dim openAIResult As String
            LogEvt PROC_NAME, lgINFO, "Attempting OpenAI call.", "Company=" & companyName ' Use lgINFO
            TraceEvt lvlINFO, PROC_NAME, "Attempting OpenAI call", "Company=" & companyName
            openAIResult = GetCompanyRecapOpenAI(companyName) ' This function logs its own success/failure

            ' Update finalRecap only if OpenAI returns a valid, non-error result
            If openAIResult <> "" And Not LCase(openAIResult) Like "error:*" Then
                finalRecap = openAIResult
                 mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "OpenAI SUCCESS, using result.", "Company=" & companyName
            Else
                 mod_DebugTraceHelpers.TraceEvt IIf(LCase(openAIResult) Like "error:*", lvlERROR, lvlWARN), PROC_NAME, "OpenAI Failed or Skipped, using default.", "Company=" & companyName & ", Result=" & openAIResult
            End If
        Else
             LogEvt PROC_NAME, lgINFO, "OpenAI call skipped (Not Maintainer or disabled).", "Company=" & companyName ' Use lgINFO
             mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "OpenAI call skipped", "Company=" & companyName
        End If

        ' 3. Add the result (Default or OpenAI) to the Memory Cache for this run
        On Error Resume Next ' Handle potential error adding to dictionary
        dictCache(companyName) = finalRecap
        If Err.Number <> 0 Then
            LogEvt PROC_NAME, lgERROR, "Error adding '" & companyName & "' to memory cache: " & Err.Description ' Use lgERROR
            mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error adding to memory cache", "Company=" & companyName & ", Err=" & Err.Description
            Err.Clear
        Else
             mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Added to memory cache", "Company=" & companyName
        End If
        On Error GoTo 0 ' Restore default error handling
    End If

    GetCompanyRecap = finalRecap
End Function

Public Sub SaveCompanyCache(wsCache As Worksheet)
    ' Purpose: Saves the in-memory company cache back to the sheet.
    Dim key As Variant, i As Long, outputArr() As Variant, saveCount As Long
    Const PROC_NAME As String = "mod_Cache.SaveCompanyCache" ' Updated PROC_NAME

    If dictCache Is Nothing Or dictCache.Count = 0 Then
        LogEvt PROC_NAME, lgINFO, "In-memory cache empty, skipping save." ' Use lgINFO
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Skipped saving empty cache"
        Exit Sub
    End If

    On Error GoTo CacheSaveError
    saveCount = dictCache.Count
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Saving cache to sheet", "Sheet=" & wsCache.Name & ", Items=" & saveCount
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
            .Range("A2").Resize(saveCount, 3).value = outputArr
            ' Format the timestamp column
            .Range("C2").Resize(saveCount, 1).NumberFormat = "m/d/yyyy h:mm AM/PM"
            ' Autofit columns after writing
             On Error Resume Next ' Autofit might fail on hidden sheets sometimes
            .Columns("A:C").AutoFit
             On Error GoTo CacheSaveError ' Restore handler
        End If
    End With
    LogEvt PROC_NAME, lgINFO, "Saved " & saveCount & " items to cache sheet." ' Use lgINFO
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Cache save complete", "ItemsSaved=" & saveCount

CacheSaveExit: ' Label for normal exit and error exit cleanup
    ' Restore application settings
    Application.EnableEvents = previousEnableEvents
    Application.Calculation = previousCalculation
    Exit Sub

CacheSaveError:
     LogEvt PROC_NAME, lgERROR, "Error saving cache to sheet '" & wsCache.Name & "': " & Err.Description ' Use lgERROR
     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error saving cache", "Sheet=" & wsCache.Name & ", Err=" & Err.Number & " - " & Err.Description
    MsgBox "Error saving company cache to sheet '" & wsCache.Name & "': " & Err.Description, vbExclamation, "Cache Save Error"
    Resume CacheSaveExit ' Attempt to restore settings even after error
End Sub

' --- Private Helper Functions ---

Private Function GetCompanyRecapOpenAI(companyName As String) As String
    ' Purpose: Calls OpenAI API to get a company summary. Includes error handling & logging.
    Dim apiKey As String, result As String, http As Object, url As String, jsonPayload As String, jsonResponse As String
    Const PROC_NAME As String = "mod_Cache.GetCompanyRecapOpenAI" ' Updated PROC_NAME
    GetCompanyRecapOpenAI = "" ' Default return value
    
    ' Check if OpenAI API calls are enabled globally
    If Not mod_Config.ENABLE_OPENAI_API_CALLS Then
        GetCompanyRecapOpenAI = "OpenAI disabled"   ' early exit
        LogEvt PROC_NAME, lgINFO, "Skipped OpenAI Call: Feature disabled globally.", "Company=" & companyName
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Skipped: Feature disabled", "Company=" & companyName
        Exit Function
    End If

    ' Double-check maintainer status (though already checked by caller)
    ' Assumes IsMaintainerUser is available (e.g., in mod_Utils or mod_Config)
    If Not IsMaintainerUser() Then ' Requires mod_Utils or similar
         LogEvt PROC_NAME, lgINFO, "Skipped OpenAI Call: Not Maintainer User.", "Company=" & companyName ' Use lgINFO
         mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Skipped: Not Maintainer", "Company=" & companyName
        Exit Function ' Should not happen if called correctly, but safe
    End If

    ' Get API Key
    apiKey = GetAPIKey() ' Assumes GetAPIKey logs its own errors/warnings
    If apiKey = "" Then
        ' GetAPIKey function should have logged the reason
        GetCompanyRecapOpenAI = "Error: API Key Not Configured" ' Return error string
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Skipped: API Key Not Found/Configured", "Company=" & companyName
        Exit Function
    End If

    On Error GoTo OpenAIErrorHandler

    ' Prepare Request
    url = OPENAI_API_URL ' Assumes constant is available from mod_Config
    Dim modelName As String: modelName = OPENAI_MODEL ' Assumes constant is available from mod_Config
    Dim systemPrompt As String, userPrompt As String
    systemPrompt = "You are an analyst summarizing medical device related companies based *only* on publicly available information. " & _
                   "Provide a *neutral*, *very concise* (1 sentence ideally, 2 max) summary of the company '" & Replace(companyName, """", "'") & "' " & _
                   "identifying its primary business sector or main product type (e.g., orthopedics, diagnostics, surgical tools, contract manufacturer). " & _
                   "If unsure or company is generic, state 'General medical device company'. Avoid speculation or marketing language."
    userPrompt = "Summarize: " & companyName

    jsonPayload = "{""model"": """ & modelName & """, ""messages"": [" & _
                  "{""role"": ""system"", ""content"": """ & JsonEscape(systemPrompt) & """}," & _
                  "{""role"": ""user"", ""content"": """ & JsonEscape(userPrompt) & """}" & _
                  "], ""temperature"": 0.3, ""max_tokens"": " & OPENAI_MAX_TOKENS & "}" ' Assumes constant is available from mod_Config

    ' Send Request
    LogEvt PROC_NAME, lgDETAIL, "Sending request...", "Company=" & companyName & ", Model=" & modelName ' Use lgDETAIL
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Sending request...", "Company=" & companyName & ", Model=" & modelName
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False ' Synchronous call
    http.setTimeouts OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS, OPENAI_TIMEOUT_MS ' Assumes constant is available from mod_Config
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send jsonPayload

    ' Process Response
    LogEvt PROC_NAME, lgDETAIL, "Response Received.", "Company=" & companyName & ", Status=" & http.Status ' Use lgDETAIL
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Response Received", "Company=" & companyName & ", Status=" & http.Status

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
                    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "OpenAI SUCCESS", "Company=" & companyName
                Else
                     result = "Error: Parse Fail (End Quote)"
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 500) ' Use lgERROR
                     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Parse Fail (End Quote)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                End If
            Else
                ' Content tag not found, check if it's an error object from OpenAI
                If InStr(1, jsonResponse, """error""", vbTextCompare) > 0 Then
                     result = "Error: API returned error object."
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
                     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "API returned error object", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                Else
                     result = "Error: Parse Fail (Start Tag)"
                     LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
                     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Parse Fail (Start Tag)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
                End If
            End If
        Else
             result = "Error: Parse Fail (No Assistant Role)"
             LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(jsonResponse, 500) ' Use lgERROR
             mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Parse Fail (No Assistant Role)", "Company=" & companyName & ", ResponseStart=" & Left(jsonResponse, 100)
        End If
    Else
        ' HTTP Error
        result = "Error: API Call Failed - Status " & http.Status & " - " & http.statusText
         LogEvt PROC_NAME, lgERROR, result, "Company=" & companyName & ", Response=" & Left(http.responseText, 500) ' Use lgERROR
         mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "API Call Failed", "Company=" & companyName & ", Status=" & http.Status & ", ResponseStart=" & Left(http.responseText, 100)
    End If

    ' Cleanup and Finalize
    Set http = Nothing
    If Len(result) > RECAP_MAX_LEN Then result = Left$(result, RECAP_MAX_LEN - 3) & "..." ' Ensure truncation fits (Assumes constant from mod_Config)
    GetCompanyRecapOpenAI = Trim(result) ' Return the parsed/error string
    Exit Function

OpenAIErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
     LogEvt PROC_NAME, lgERROR, "VBA Exception during OpenAI Call: " & errDesc, "Company=" & companyName ' Use lgERROR
     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "VBA Exception", "Company=" & companyName & ", Err=" & Err.Number & " - " & errDesc
    GetCompanyRecapOpenAI = "Error: VBA Exception - " & errDesc ' Return VBA error string
    If Not http Is Nothing Then Set http = Nothing ' Clean up object on error
End Function

Private Function GetAPIKey() As String
    ' Purpose: Reads the OpenAI API key from a specified file path.
    Dim fso As Object, ts As Object, keyPath As String, WshShell As Object, fileContent As String: fileContent = ""
    Const PROC_NAME As String = "mod_Cache.GetAPIKey" ' Updated PROC_NAME
    On Error GoTo KeyError

    ' Expand environment variables in the path
    Set WshShell = CreateObject("WScript.Shell")
    keyPath = WshShell.ExpandEnvironmentStrings(API_KEY_FILE_PATH) ' Assumes constant from mod_Config
    Set WshShell = Nothing
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Resolved API Key Path", keyPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(keyPath) Then
        Set ts = fso.OpenTextFile(keyPath, 1) ' ForReading
        If Not ts.AtEndOfStream Then fileContent = ts.ReadAll
        ts.Close
        If Len(Trim(fileContent)) > 0 Then
             LogEvt PROC_NAME, lgDETAIL, "API Key read successfully." ' Use lgDETAIL
             mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "API Key read successfully"
        Else
             LogEvt PROC_NAME, lgWARN, "API Key file exists but is empty.", "Path=" & keyPath ' Use lgWARN
             mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "API Key file empty", "Path=" & keyPath
        End If
    Else
         LogEvt PROC_NAME, lgWARN, "API Key file not found.", "Path=" & keyPath ' Use lgWARN
         mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "API Key file not found", "Path=" & keyPath
        Debug.Print Time & " - WARNING: API Key file not found at specified path: " & keyPath
    End If
    GoTo KeyExit

KeyError:
     LogEvt PROC_NAME, lgERROR, "Error reading API Key from '" & keyPath & "': " & Err.Description ' Use lgERROR
     mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error reading API Key file", "Path=" & keyPath & ", Err=" & Err.Number & " - " & Err.Description
    Debug.Print Time & " - ERROR reading API Key from '" & keyPath & "': " & Err.Description

KeyExit:
    GetAPIKey = Trim(fileContent)
    ' Clean up objects
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: Set ts = Nothing
    If Not fso Is Nothing Then Set fso = Nothing
    On Error GoTo 0 ' Restore default error handling
End Function

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

Private Function JsonUnescape(strInput As String) As String
    ' Purpose: Unescapes characters in a string retrieved from a JSON payload.
    strInput = Replace(strInput, "\n", vbCrLf) ' Convert \n back to CRLF
    strInput = Replace(strInput, "\t", vbTab)  ' Convert \t back to Tab
    strInput = Replace(strInput, "\""", """") ' Unescape double quotes
    strInput = Replace(strInput, "\\", "\")   ' Unescape backslashes LAST
    JsonUnescape = strInput
End Function
