Attribute VB_Name = "mod_DataIO_Fixed"
Option Explicit

' ==========================================================================
' Module      : mod_DataIO_Fixed
' Author      : Cline (AI Assistant)
' Date        : 2025-05-07
' Version     : 1.0.0
' ==========================================================================
' Description : Fixed version of functions from mod_DataIO module with 
'               connection cleanup bug fixes.
'
' Key Feature : CleanupDuplicateConnections() - Rewritten with a safer
'               approach to avoid Error 9 "Subscript out of range" when
'               deleting connections during iteration.
'
' Dependencies: - mod_Logger: For logging operations.
'               - mod_DebugTraceHelpers: For detailed debug tracing.
'
' Revision History:
' --------------------------------------------------------------------------
' Date        Author          Description
' ----------- --------------- ----------------------------------------------
' 2025-05-07  Cline (AI)      - Initial version with fixed CleanupDuplicateConnections
' ==========================================================================

' --- Constants ---
Private Const MODULE_NAME As String = "mod_DataIO_Fixed"

' ==========================================================================
' ===                   CONNECTION MANAGEMENT                           ===
' ==========================================================================

Public Sub CleanupDuplicateConnections()
    ' Purpose: Safely cleans up duplicate or orphaned connections
    ' This version uses a two-phase approach to avoid collection modification issues:
    ' 1. First identify all connections to delete
    ' 2. Then delete them (in reverse order) after identification is complete
    
    Const PROC_NAME As String = "CleanupDuplicateConnections"
    
    ' Variables for connection processing
    Dim c As WorkbookConnection
    Dim connectionsToDelete As Collection
    Dim i As Integer, j As Integer
    Dim connectionList As String
    Dim connCount As Integer
    
    ' Initialize collection to store connections for deletion
    Set connectionsToDelete = New Collection
    
    ' Log starting connection count
    connCount = ThisWorkbook.Connections.Count
    LogEvt PROC_NAME, lgINFO, "Starting connection cleanup. Initial count: " & connCount
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Starting cleanup", "Count=" & connCount
    
    ' Skip processing if no connections exist
    If connCount = 0 Then
        LogEvt PROC_NAME, lgINFO, "No connections found - nothing to clean up."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "No connections found"
        Exit Sub
    End If
    
    ' ===== PHASE 1: Identify connections to delete =====
    On Error Resume Next
    For i = 1 To ThisWorkbook.Connections.Count
        Set c = ThisWorkbook.Connections(i)
        
        If Not c Is Nothing Then
            ' Log each connection for debugging
            LogEvt PROC_NAME, lgDETAIL, "Connection " & i & ": " & c.Name
            
            ' Check for duplicate or stale connections based on naming patterns
            ' and other criteria
            
            ' Example: Check for connections with duplicate prefix/naming patterns
            If InStr(1, c.Name, "Query - pgGet510kData") > 0 Then
                ' This is likely a duplicate PQ connection
                LogEvt PROC_NAME, lgDETAIL, "Marking for deletion: " & c.Name & " (duplicate prefix)"
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Marked for deletion", "Name=" & c.Name & ", Reason=duplicate prefix"
                connectionsToDelete.Add i, CStr(i)
            ElseIf InStr(1, c.Name, "_xlnm.") > 0 Then
                ' This is an Excel-generated connection that might be stale
                LogEvt PROC_NAME, lgDETAIL, "Marking for deletion: " & c.Name & " (Excel-generated)"
                mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Marked for deletion", "Name=" & c.Name & ", Reason=Excel-generated"
                connectionsToDelete.Add i, CStr(i)
            End If
            
            ' Add additional criteria as needed based on your specific requirements
            ' For example, checking connection state, target, etc.
        End If
    Next i
    On Error GoTo 0
    
    ' ===== PHASE 2: Delete identified connections (in reverse order) =====
    If connectionsToDelete.Count > 0 Then
        LogEvt PROC_NAME, lgINFO, "Found " & connectionsToDelete.Count & " connections to remove."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Connections to remove", "Count=" & connectionsToDelete.Count
        
        ' Process in REVERSE order to avoid index shifting problems
        For i = connectionsToDelete.Count To 1 Step -1
            On Error Resume Next
            j = connectionsToDelete(i) ' Get the stored index
            
            ' Get connection name before deletion (for logging)
            Dim connName As String
            connName = "Unknown"
            If j <= ThisWorkbook.Connections.Count Then
                If Not ThisWorkbook.Connections(j) Is Nothing Then
                    connName = ThisWorkbook.Connections(j).Name
                End If
            End If
            
            ' Delete the connection
            ThisWorkbook.Connections(j).Delete
            
            ' Log result
            If Err.Number = 0 Then
                LogEvt PROC_NAME, lgINFO, "Successfully removed connection: " & connName
                mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Removed connection", "Name=" & connName
            Else
                LogEvt PROC_NAME, lgWARN, "Failed to remove connection " & connName & ". Error: " & Err.Number & " - " & Err.Description
                mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Failed to remove connection", "Name=" & connName & ", Err=" & Err.Number
            End If
            Err.Clear
            On Error GoTo 0
        Next i
    Else
        LogEvt PROC_NAME, lgINFO, "No connections identified for removal."
        mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "No connections to remove"
    End If
    
    ' Log final connection count
    LogEvt PROC_NAME, lgINFO, "Cleanup complete. Final count: " & ThisWorkbook.Connections.Count
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Cleanup complete", "FinalCount=" & ThisWorkbook.Connections.Count
    
    ' Clean up
    Set connectionsToDelete = Nothing
End Sub
