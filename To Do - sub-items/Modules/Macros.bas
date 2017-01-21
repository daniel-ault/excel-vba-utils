Attribute VB_Name = "Macros"


Public Sub MoveToNextStage()
    Dim sht As Worksheet
    Set sht = ActiveSheet
    
    Dim CurrentSheet As String
    CurrentSheet = sht.Name
    
    'MsgBox GetNextStage(sht.Name)
    
    Dim row As Integer
    row = Selection.row
    
    Call MoveRow(Selection.row, GetNextStage(CurrentSheet))
End Sub

Public Function GetNextStage(CurrentStage As String)
    Dim Settings As Worksheet
    Set Settings = Worksheets("Settings")
    
    'Dim stagesTop As Integer
    stagesTop = 2
    stagesBottom = Settings.Range("K" & stagesTop).End(xlDown).row
    
    Dim stagesRange As Range
    Set stagesRange = Settings.Range("K" & stagesTop, "K" & stagesBottom)
    
    Dim blnIsNext As Boolean
    blnIsNext = False
    
    For Each cell In stagesRange
        If blnIsNext = True Then
            GetNextStage = cell.Value
            Exit Function
        End If
        If CurrentStage = cell.Value Then blnIsNext = True
    Next
End Function

Public Sub MoveToILOList()
    Dim row As Integer
    row = Selection.row
    
    Dim DestSheet As String
    DestSheet = "Push to ILO"
    
    Call MoveRow(Selection.row, "Push to ILO")
End Sub


Public Sub MoveRow(row As Integer, DestSheet As String)
    Dim rngStr As String
    rngStr = "B" & row & ":G" & row
    
    LastRow = Worksheets(DestSheet).Range("B1").End(xlDown).row
    Dim EmptyRow As Integer
    EmptyRow = LastRow + 1
    
    ActiveSheet.Range(rngStr).Copy _
        Worksheets(DestSheet).Range("B" & EmptyRow)
    
    'set complete date on current row
    'ActiveSheet.Range("G" & Row).value = Date
    Call FinishRow
End Sub


Public Sub FinishRow()
    For Each r In Selection.Rows
         ActiveSheet.Range("G" & r.row()).Value = Date
    Next
    ActiveSheet.AutoFilter.ApplyFilter
End Sub


Public Sub RefreshFilters()
    ActiveSheet.AutoFilter.ApplyFilter
    ActiveSheet.Range("I2").Calculate
End Sub

'Public Sub RefreshFilters(sheet As String)
'    Worksheets(sheet).AutoFilter.ApplyFilter
'End Sub

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
                .Cells(16 + x, 11) = "Row " & i & " is visible"
                x = x + 1
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


