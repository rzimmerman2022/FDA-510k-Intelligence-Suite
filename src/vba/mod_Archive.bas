' ==========================================================================
' Module      : mod_Archive
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module manages the archiving process for the processed
'               510(k) data. When triggered (typically once per month), it
'               creates a static copy of the main data sheet, naming it
'               according to the processed month (e.g., "Apr-2025"). It
'               converts the table object on the newly created archive sheet
'               into a static range to preserve the data snapshot and prevent
'               accidental refreshes. It also handles the cleanup of potential
'               duplicate Power Query connections that can be created when
'               copying worksheets containing query tables.
'
' Key Function: ArchiveIfNeeded(tblData As ListObject, archiveSheetName As String) As Boolean
'               - Performs the sheet copy, rename, table unlist, and connection
'                 cleanup operations. Called conditionally by the main
'                 processing routine in mod_510k_Processor.
'
' Dependencies: - mod_Logger: For logging archiving steps and errors.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'               - mod_DataIO: Specifically calls CleanupDuplicateConnections
'                 to manage Power Query connections after sheet copy.
'               - Assumes the main data table object (tblData) and the
'                 desired archive sheet name are passed correctly.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit
Attribute VB_Name = "mod_Archive"

Public Function ArchiveIfNeeded(tblData As ListObject, archiveSheetName As String) As Boolean
    ' Purpose: Creates an archive copy of the data sheet if needed.
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

    LogEvt PROC_NAME, lgINFO, "Archiving data sheet '" & wsData.Name & "' to '" & archiveSheetName & "'."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Start archiving", "SourceSheet=" & wsData.Name & ", ArchiveSheet=" & archiveSheetName

    ' --- Disable alerts and screen updates for smoother copy ---
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' --- Copy the current data sheet ---
    wsData.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsArchive = ActiveSheet ' The newly copied sheet becomes active

    ' --- Rename the copied sheet ---
    On Error Resume Next ' Handle error if sheet name already exists (shouldn't happen if mustArchive was true)
    wsArchive.Name = archiveSheetName
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Failed to rename copied sheet to '" & archiveSheetName & "'. Error: " & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Failed to rename archive sheet", "SheetName=" & archiveSheetName & ", Err=" & Err.Description
        ' Attempt to delete the failed copy
        On Error Resume Next
        wsArchive.Delete
        On Error GoTo ArchiveError ' Restore handler
        GoTo ArchiveCleanup ' Go to cleanup without setting success flag
    End If
    On Error GoTo ArchiveError ' Restore handler

    ' --- Remove Table functionality from the archive sheet (convert to range) ---
    Dim tblArchive As ListObject
    On Error Resume Next ' Handle error finding the table on the new sheet
    Set tblArchive = wsArchive.ListObjects(1) ' Assumes one table per sheet
    If Err.Number = 0 And Not tblArchive Is Nothing Then
        tblArchive.Unlist ' Convert table to range
        LogEvt PROC_NAME, lgDETAIL, "Converted table to range on archive sheet.", "Sheet=" & archiveSheetName
        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Converted table to range", "Sheet=" & archiveSheetName
    Else
        LogEvt PROC_NAME, lgWARN, "Could not find or unlist table on archive sheet '" & archiveSheetName & "'.", "Err=" & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Could not unlist table on archive", "Sheet=" & archiveSheetName & ", Err=" & Err.Description
        Err.Clear ' Clear error and continue
    End If
    On Error GoTo ArchiveError ' Restore handler

    ' --- Optional: Remove formulas, keep values (if needed) ---
    ' wsArchive.UsedRange.Value = wsArchive.UsedRange.Value

    ' --- Clean up duplicate Power Query connections created by the copy ---
    Call mod_DataIO.CleanupDuplicateConnections ' Assumes this is safe to call after copy

    ArchiveIfNeeded = True ' Success
    LogEvt PROC_NAME, lgINFO, "Successfully archived sheet to '" & archiveSheetName & "'."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Archiving successful", "ArchiveSheet=" & archiveSheetName

ArchiveCleanup:
    ' --- Restore settings ---
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Set wsArchive = Nothing
    Set tblArchive = Nothing
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
