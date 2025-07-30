' ==========================================================================
' Module      : mod_Archive
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module manages the archiving process for the processed
'               510(k) data. When triggered (typically once per month), it
'               creates a static copy of the main data as a values-only range
'               on a new worksheet, naming it according to the processed month
'               (e.g., "Apr-2025"). Unlike previous implementations, it now
'               preserves the original table object on the data sheet, enabling
'               subsequent Power Query refreshes to succeed. This implementation
'               avoids the Error #1004 (Application-defined or object-defined error)
'               that occurred when the original QueryTable was unlinked from its
'               ListObject.
'
' Key Function: ArchiveIfNeeded(tblData As ListObject, archiveSheetName As String) As Boolean
'               - Creates a values-only copy of the table data on a new sheet,
'                 while preserving the original table structure and QueryTable
'                 connection. Called conditionally by the main processing routine
'                 in mod_510k_Processor.
'
' Dependencies: - mod_Logger: For logging archiving steps and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - Assumes the main data table object (tblData) and the
'                 desired archive sheet name are passed correctly.
'               NOTE: This module no longer depends on mod_DataIO's
'               CleanupDuplicateConnections function since we now create
'               a values-only copy rather than copying the entire sheet.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-05-08  Cline (AI)      - Modified ArchiveIfNeeded to preserve the original table
'                              structure rather than converting to a static range.
'                              This fixes Error #1004 during PQ refresh after archiving.
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit

Public Function ArchiveIfNeeded(tblData As ListObject, archiveSheetName As String) As Boolean
    ' Purpose: Creates an archive copy of the data sheet if needed.
    '          IMPORTANT: This function now preserves the original table structure
    '          and creates a static values-only copy in a new sheet, rather than
    '          unlisting the original table. This ensures that subsequent
    '          Power Query refreshes will continue to work properly.
    ' Returns: True if successful or not needed, False on critical error during archive.
    Const PROC_NAME As String = "mod_Archive.ArchiveIfNeeded"
    ArchiveIfNeeded = False ' Default to failure

    If tblData Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Invalid argument (Table is Nothing). Cannot archive."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Invalid argument", "TableIsNothing=True"
        Exit Function
    End If

    Dim wsData As Worksheet: Set wsData = tblData.Parent
    Dim wsArchive As Worksheet

    On Error GoTo ArchiveError

    LogEvt PROC_NAME, lgINFO, "Archiving data from '" & wsData.Name & "' to new sheet '" & archiveSheetName & "'."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start archiving", "SourceSheet=" & wsData.Name & ", ArchiveSheet=" & archiveSheetName

    ' --- Disable alerts and screen updates for smoother operation ---
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' --- Create a new sheet for the archive ---
    On Error Resume Next ' Handle error if sheet creation fails
    Set wsArchive = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Failed to add new sheet for archive. Error: " & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Failed to add new sheet", "Err=" & Err.Description
        On Error GoTo ArchiveError ' Restore handler
        GoTo ArchiveCleanup ' Go to cleanup without setting success flag
    End If
    On Error GoTo ArchiveError ' Restore handler
    
    ' --- Rename the new sheet ---
    On Error Resume Next ' Handle error if sheet rename fails
    wsArchive.Name = archiveSheetName
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Failed to rename new sheet to '" & archiveSheetName & "'. Error: " & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Failed to rename archive sheet", "SheetName=" & archiveSheetName & ", Err=" & Err.Description
        ' Attempt to delete the failed sheet
        On Error Resume Next
        wsArchive.Delete
        On Error GoTo ArchiveError ' Restore handler
        GoTo ArchiveCleanup ' Go to cleanup without setting success flag
    End If
    On Error GoTo ArchiveError ' Restore handler
    
    ' --- Copy the data as values to the new sheet ---
    tblData.Range.Copy
    wsArchive.Range("A1").PasteSpecial xlPasteValues
    wsArchive.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    LogEvt PROC_NAME, lgDETAIL, "Copied table data as values to archive sheet.", "Sheet=" & archiveSheetName
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Copied data as values", "Sheet=" & archiveSheetName

    ' --- Format the archive sheet for better readability (optional) ---
    On Error Resume Next ' Ignore formatting errors
    wsArchive.UsedRange.Columns.AutoFit
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgDETAIL, "Minor error during column auto-fit: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ArchiveError ' Restore handler

    ' --- No need to clean up connections since we're creating a new sheet with values only ---
    
    ArchiveIfNeeded = True ' Success
    LogEvt PROC_NAME, lgINFO, "Successfully archived data to '" & archiveSheetName & "' while preserving original table."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Archiving successful", "ArchiveSheet=" & archiveSheetName

ArchiveCleanup:
    ' --- Restore settings ---
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Set wsArchive = Nothing
    Exit Function

ArchiveError:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    LogEvt PROC_NAME, lgERROR, "Error during archiving process for '" & archiveSheetName & "'. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error during archiving", "ArchiveSheet=" & archiveSheetName & ", Err=" & errNum & " - " & errDesc
    MsgBox "An error occurred during the archiving process: " & vbCrLf & errDesc, vbCritical, "Archiving Error"
    ' ArchiveIfNeeded remains False
    Resume ArchiveCleanup ' Go to cleanup section
End Function
