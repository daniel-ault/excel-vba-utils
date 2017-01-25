'Private Sub Worksheet_Activate()
'    'Call RefreshFilters
'    'Call CopyFilteredRange
'End Sub

Private Sub GetFilteredData()
    Dim strCategory As String
    strCategory = "Excel To-Do List"
    
    Dim shtMasterSheet As Worksheet
    Set shtMasterSheet = Worksheets("To Do List")
    
    With shtMasterSheet
        .AutoFilterMode = False
        With .Range("B2:G999")
            .AutoFilter Field:=1, Criteria1:=strCategory
            .SpecialCells(xlCellTypeVisible).Copy Destination:=Worksheets("To Do List (2)").Range("A2")
        End With
        .AutoFilterMode = True
    End With
End Sub


Private Sub CopyFilteredRange()
    Dim row As Range
    Dim NewRange As Range
    Dim currentRow As Integer
    currentRow = 2
    
    Dim category As String
    category = "Excel To-Do List"
    
    'TODO clear range first?
    'Worksheets("To Do List (2)").Range("Table65").Clear
    
    'MsgBox Worksheets("To Do List").Range("Table6").Rows.Count
    Worksheets("To Do List").Range("Table6").Copy
    Worksheets("To Do List (2)").Range("Table65").PasteSpecial

    'Worksheets("To Do List (2)").Range("Table65").value = Worksheets("To Do List").Range("Table6").value

    Dim tbl As Range
    Set tbl = Worksheets("To Do List (2)").Range("Table65")
    
    Dim rowCount As Integer
    rowCount = tbl.Rows.count
    
    Dim i As Integer
    'For i = rowCount To 1
        
    'Next i
    
    For Each row In tbl.Rows
        'MsgBox row.Cells(2).value
        If row.Cells(2).Value <> category Then
            row.Delete
        End If
        
    Next row
    
'    For Each row In Worksheets("To Do List").Range("Table6").Rows
'        'MsgBox row.Cells(2).value
'        If row.Cells(2) = category Then
'            MsgBox row.Address
'            'row.Copy
'            'Worksheets("To Do List (2)").Range("A" & currentRow).PasteSpecial
'            'MsgBox row.Address
'            currentRow = currentRow + 1
'        End If
'    Next row
    
    Application.CutCopyMode = False
End Sub
