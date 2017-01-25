Public Sub NumberRows()
    Dim col As Integer
    Dim row As Integer
    col = 1
    row = 2
    
    For row = 2 To 301
        Cells(row, col).Value = row - 1
    Next row
End Sub

Public Sub FilteredRows()
    For i = 1 To 50
        With Excel.ThisWorkbook.ActiveSheet
            If .Rows(i).EntireRow.Hidden Then
            Else
                .Cells(16 + X, 11) = "Row " & i & " is visible"
                X = X + 1
            End If
        End With
    Next i
End Sub

Public Sub LastColumn() 'Optional ByVal Row As Integer = 0)
    Dim row As Integer
    row = Selection.row
    
    Set sheet = Excel.ThisWorkbook.ActiveSheet
    Last = sheet.Range("A" & row).CurrentRegion.Columns.count
    
    MsgBox Last
End Sub

