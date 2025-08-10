' ==========================================================================
' Module      : mod_Debug
' Author      : [Original Author - Unknown]
' Date        : [Original Date - Unknown]
' Maintainer  : Cline (AI Assistant)
' Version     : See mod_Config.VERSION_INFO
' ==========================================================================
' Description : This module provides basic helper subroutines primarily
'               intended for debugging during development. It includes
'               functionality to dump table headers to the Immediate Window
'               and a simple tracing mechanism that writes messages to a
'               dedicated "DebugTrace" worksheet.
'
'               NOTE: The DebugTrace subroutine appears to be a separate,
'               simpler logging mechanism compared to the more comprehensive
'               TraceEvt system implemented in mod_DebugTraceHelpers.
'               Consider consolidating tracing efforts into mod_DebugTraceHelpers
'               in future refactoring.
'
' Key Functions:
'               - DumpHeaders: Prints the column names and indices of the
'                 first table found on the "CurrentMonthData" sheet to the
'                 VBA Immediate Window. Useful for verifying column order.
'               - DebugTrace: Writes a timestamped message with a tag and
'                 optional caller info to the "DebugTrace" sheet, creating
'                 the sheet if it doesn't exist.
'
' Dependencies: - Assumes a worksheet named "CurrentMonthData" exists for
'                 DumpHeaders.
'               - Interacts directly with Workbook sheets for DebugTrace.
'               - Does not appear to have dependencies on other custom modules.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-04-30  Cline (AI)      - Added detailed module header comment block.
' [Previous dates/authors/changes unknown]
' ==========================================================================
Option Explicit

Sub DumpHeaders()
    ' Purpose: Prints the column order and names of the first table on the active sheet to the Immediate Window.
    Dim lo As ListObject
    Dim c As Range
    Dim ws As Worksheet

    ' Ensure the target sheet is active or specify it directly
    On Error Resume Next ' Handle error if sheet doesn't exist or no table
    Set ws = ThisWorkbook.Sheets("CurrentMonthData") ' Specify the sheet name directly
    If ws Is Nothing Then
         Debug.Print "ERROR (DumpHeaders): Sheet 'CurrentMonthData' not found."
         Exit Sub
    End If
    If ws.ListObjects.Count = 0 Then
         Debug.Print "ERROR (DumpHeaders): No ListObject (table) found on sheet 'CurrentMonthData'."
         Exit Sub
    End If
    Set lo = ws.ListObjects(1) ' Assumes the first table is the target
    On Error GoTo 0 ' Restore error handling

    Debug.Print "=== Header order from DumpHeaders (Sheet: " & ws.Name & ", Table: " & lo.Name & ") ==="
    If lo.HeaderRowRange Is Nothing Then
         Debug.Print "ERROR (DumpHeaders): Table '" & lo.Name & "' does not have a header row range."
         Exit Sub
    End If

    For Each c In lo.HeaderRowRange.Cells
        Debug.Print "Col " & c.Column & ": ", c.value
    Next c
    Debug.Print "========================================="

End Sub
