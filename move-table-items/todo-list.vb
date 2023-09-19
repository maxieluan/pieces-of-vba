
Sub MoveCOmpletedTasks()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim tblSource As ListObject
    Dim tblRow As ListRow
    Dim newRow As Range
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    
    Set wsSource = ThisWorkbook.Sheets("待办事项列表")
    Set wsDestination = ThisWorkbook.Sheets("History")
    
    
    Set tblSource = wsSource.ListObjects("待办事项列表")
    
    For i = tblSource.ListRows.Count - 1 To 1 Step -1
        Set tblRow = tblSource.ListRows(i)
    
        If tblRow.Range(1, 2).Value >= 1 Then
            Set newRow = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Offset(1, 0)
            For j = 1 To tblRow.Range.Columns.Count
                newRow.Offset(0, j - 1).Value = tblRow.Range(1, j).Value
            Next j
            tblRow.Delete
        End If
    Next i
End Sub
