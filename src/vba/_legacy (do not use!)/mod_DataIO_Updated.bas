' If the base exists, delete all numbered variants
        If baseExists Then
            Dim baseName As String: baseName = baseConnectionNames(i)
            LogEvt PROC_NAME, lgINFO, "Found primary connection '" & baseName & "'. Will preserve this and delete duplicates."
            mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Found primary connection", "Base=" & baseName
            
            ' Loop through all connections to find numbered variants
            Dim c2 As WorkbookConnection
            For Each c2 In ThisWorkbook.Connections
                ' Check if it's a numbered variant
                If c2.Name <> baseName And c2.Name Like baseName & " (*)" Then
                    LogEvt PROC_NAME, lgDETAIL, "Deleting duplicate connection: " & c2.Name
                    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting duplicate connection", c2.Name
                    c2.Delete
                    deletedCount = deletedCount + 1
                End If
            Next c2
        End If
    Next i
    
    ' Second pass: If we still don't have a primary connection but have duplicates with numbers,
    ' keep the lowest numbered one and delete the rest
    Dim lowestNumberedConn As WorkbookConnection
    Dim lowestNumber As Integer: lowestNumber = 9999
    
    For i = LBound(baseConnectionNames) To UBound(baseConnectionNames)
        Dim baseName2 As String: baseName2 = baseConnectionNames(i)
        
        ' Check again if the base exists (it might after the first pass)
        Set c = Nothing
        Set c = ThisWorkbook.Connections(baseName2)
        If c Is Nothing Then
            ' Base doesn't exist, look for numbered variants
            LogEvt PROC_NAME, lgDETAIL, "Primary connection '" & baseName2 & "' not found. Looking for lowest numbered variant."
            mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Primary connection not found", "Base=" & baseName2
            
            ' Find the lowest numbered variant
            Set lowestNumberedConn = Nothing
            lowestNumber = 9999
            
            For Each c In ThisWorkbook.Connections
                If c.Name Like baseName2 & " (*)" Then
                    ' Extract the number from the connection name
                    Dim connNumber As String
                    Dim startPos As Integer, endPos As Integer
                    
                    startPos = InStr(c.Name, "(") + 1
                    endPos = InStr(c.Name, ")")
                    
                    If startPos > 1 And endPos > startPos Then
                        connNumber = Mid(c.Name, startPos, endPos - startPos)
                        Dim num As Integer
                        
                        If IsNumeric(connNumber) Then
                            num = CInt(connNumber)
                            If num < lowestNumber Then
                                lowestNumber = num
                                Set lowestNumberedConn = c
                            End If
                        End If
                    End If
                End If
            Next c
            
            ' If we found a lowest numbered connection, rename it to the base name
            If Not lowestNumberedConn Is Nothing Then
                LogEvt PROC_NAME, lgINFO, "Renaming lowest numbered connection '" & lowestNumberedConn.Name & "' to '" & baseName2 & "'"
                mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Renaming lowest numbered connection", "From=" & lowestNumberedConn.Name & ", To=" & baseName2
                
                Dim oldName As String: oldName = lowestNumberedConn.Name
                On Error Resume Next
                lowestNumberedConn.Name = baseName2
                
                If Err.Number <> 0 Then
                    LogEvt PROC_NAME, lgWARN, "Could not rename connection '" & oldName & "' to '" & baseName2 & "'. Error: " & Err.Description
                    mod_DebugTraceHelpers.TraceEvt lvlWARN, PROC_NAME, "Failed to rename connection", "Error=" & Err.Description
                    Err.Clear
                End If
                
                ' Now delete all other numbered variants of this base
                For Each c In ThisWorkbook.Connections
                    If c.Name <> baseName2 And c.Name Like baseName2 & " (*)" And c.Name <> oldName Then
                        LogEvt PROC_NAME, lgDETAIL, "Deleting additional duplicate connection: " & c.Name
                        mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Deleting additional duplicate connection", c.Name
                        c.Delete
                        deletedCount = deletedCount + 1
                    End If
                Next c
            End If
        End If
    Next i
    
    ' Final check - list remaining connections
    connectionList = ""
    For Each c In ThisWorkbook.Connections
        connectionList = connectionList & c.Name & ", "
    Next c
    If Len(connectionList) > 2 Then connectionList = Left(connectionList, Len(connectionList) - 2)
    
    LogEvt PROC_NAME, lgINFO, "Cleanup complete. Deleted " & deletedCount & " duplicate connections. Remaining connections: " & connectionList
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Cleanup complete", "DeletedCount=" & deletedCount & ", Remaining=" & connectionList
    
    On Error GoTo 0 ' Restore default error handling
    Set c = Nothing
    Set lowestNumberedConn = Nothing
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Phase: Enhanced Cleanup Connections End"
End Sub

Public Function ArrayToTable(dataArr As Variant, targetTable As ListObject) As Boolean
    ' Purpose: Writes a 2D data array back to the DataBodyRange of a target ListObject.
    ' Returns: True if successful, False otherwise.
    Const PROC_NAME As String = "ArrayToTable"
    ArrayToTable = False ' Default to failure

    If targetTable Is Nothing Then
        LogEvt PROC_NAME, lgERROR, "Target table object is Nothing. Cannot write array."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Target table object is Nothing"
        Exit Function
    End If

    If Not IsArray(dataArr) Then
        LogEvt PROC_NAME, lgERROR, "Input dataArr is not a valid array. Cannot write."
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Input dataArr is not an array", "Table=" & targetTable.Name
        Exit Function
    End If

    On Error GoTo WriteErrorHandler

    Dim numRows As Long, numCols As Long
    On Error Resume Next ' Check array bounds safely
    numRows = UBound(dataArr, 1) - LBound(dataArr, 1) + 1
    numCols = UBound(dataArr, 2) - LBound(dataArr, 2) + 1
    If Err.Number <> 0 Then
        LogEvt PROC_NAME, lgERROR, "Error getting bounds of dataArr. Cannot write.", "Table=" & targetTable.Name & ", Err=" & Err.Description
        mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error getting array bounds", "Table=" & targetTable.Name & ", Err=" & Err.Description
        Err.Clear
        On Error GoTo WriteErrorHandler ' Restore handler
        Exit Function
    End If
    On Error GoTo WriteErrorHandler ' Restore handler

    LogEvt PROC_NAME, lgDETAIL, "Attempting to write array (" & numRows & "x" & numCols & ") to table '" & targetTable.Name & "'."
    mod_DebugTraceHelpers.TraceEvt lvlDET, PROC_NAME, "Start writing array to table", "Table=" & targetTable.Name & ", Size=" & numRows & "x" & numCols

    ' --- Resize table if necessary (optional, but safer) ---
    ' Clear existing data first to avoid issues if new array is smaller
    If targetTable.ListRows.Count > 0 Then
        targetTable.DataBodyRange.ClearContents ' Changed from .Delete to .ClearContents
    End If
    ' Resize based on array dimensions (if table allows resizing)
    ' Note: Resizing might fail if table is linked externally in certain ways.
    ' Consider adding more robust resizing logic if needed.
    ' For now, assume direct write is sufficient if dimensions match or table auto-expands.

    ' --- Write the array ---
    targetTable.DataBodyRange.Resize(numRows, numCols).value = dataArr

    ArrayToTable = True ' Success
    LogEvt PROC_NAME, lgINFO, "Successfully wrote array to table '" & targetTable.Name & "'."
    mod_DebugTraceHelpers.TraceEvt lvlINFO, PROC_NAME, "Array write successful", "Table=" & targetTable.Name
    Exit Function

WriteErrorHandler:
    Dim errDesc As String: errDesc = Err.Description
    Dim errNum As Long: errNum = Err.Number
    ArrayToTable = False ' Ensure False is returned
    LogEvt PROC_NAME, lgERROR, "Error writing array to table '" & targetTable.Name & "'. Error #" & errNum & ": " & errDesc
    mod_DebugTraceHelpers.TraceEvt lvlERROR, PROC_NAME, "Error writing array to table", "Table='" & targetTable.Name & "', Err=" & errNum & " - " & errDesc
    MsgBox "Error writing data back to table '" & targetTable.Name & "': " & vbCrLf & errDesc, vbExclamation, "Write Error"
    ' Exit Function ' Exit implicitly after error handler
End Function
