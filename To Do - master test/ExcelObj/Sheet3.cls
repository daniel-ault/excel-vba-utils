Public Sub NumberRows()
    Dim col As Integer
    Dim row As Integer
    col = 1
    row = 2
    
    For row = 2 To 301
        Cells(row, col).Value = row - 1
    Next row
End Sub
