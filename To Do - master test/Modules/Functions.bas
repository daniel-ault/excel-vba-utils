Attribute VB_Name = "Functions"
Public Function GetColorNum(Optional cell As Range = Nothing)
    Dim C As Long
    
    If cell Is Nothing Then
        'Cell = Application.Caller
        C = Application.Caller.Interior.Color
    ElseIf cell.count > 1 Then
        MsgBox "Error, only specify one cell."
    Else
        C = cell.Interior.Color
    End If
    
    'Dim CellColor As Variant
    'CellColor = Cell.Interior.Color
    
    Dim r, G, B As Long
    
    r = C Mod 256
    G = C \ 256 Mod 256
    B = C = 65536 Mod 256
    
    GetColorNum = (CStr(r) & " " & CStr(G) & " " & CStr(B))
    'GetColorNum = (0.299 * R + 0.587 * G + 0.114 * B)
End Function


Public Function CategoryColor(ByVal category As String)
    For Each cell In Worksheets("Settings").Range("B2", "B6")
        If category = cell.Value Then
            CategoryColor = cell.Offset(0, -1).Value
            Exit Function
        End If
    Next

    CategoryColor = "blue"
End Function



Public Function IsRowFiltered(Optional ByVal row As Integer = 0)
    If row <= 0 Then
        row = Application.Caller.row
    End If

    IsRowFiltered = Application.ActiveSheet.Rows(row).EntireRow.Hidden
End Function
'
'
'
Public Function VisibleRowNum(Optional ByVal row As Integer = 0)
'    If Row <= 0 Then
'        Row = Application.Caller.Row
'    End If
'
'    Dim VisibleCount As Integer
'    VisibleCount = 0
'
'    For i = 2 To Row
'        If Not IsRowFiltered(i) Then
'            VisibleCount = VisibleCount + 1
'        End If
'    Next i
'
'    VisibleRowNum = VisibleCount

    'VisibleRowNum = Worksheets("To Do List").Cells(Row, 9).Value

    'VisibleRowNum = Application.ActiveSheet.Cells(row, 9).Value
    VisibleRowNum = row
End Function



Public Function IsRowDark(Optional ByVal row As Integer = 0)
    If row <= 0 Then
        row = Application.Caller.row
    End If

    IsRowDark = (VisibleRowNum(row) Mod 2) = 1
End Function



Public Function IsRowLight(Optional ByVal row As Integer = 0)
    If row <= 0 Then
        row = Application.Caller.row
    End If

    IsRowLight = (VisibleRowNum(row) Mod 2) = 0
End Function



Public Function GetCellColor(Optional cell As Range = Nothing)
    If cell Is Nothing Then
        GetCellColor = Application.Caller.Interior.Color
    ElseIf cell.count > 1 Then
        MsgBox "Error, only specify one cell."
    Else
        GetCellColor = Application.Caller.Interior.Color
    End If

    GetCellColor = cell.Interior.Color
End Function

Public Function GetCellColorString(Optional cell As Range = Nothing)
    Dim Color As Variant
    If cell Is Nothing Then
        Color = Application.Caller.Interior.Color
    ElseIf cell.count > 1 Then
        MsgBox "Error, only specify one cell."
    Else
        Color = Application.Caller.Interior.Color
    End If
    
    For Each cell In Worksheets("Settings").Range("F2", "F6")
        If cell.Interior.Color = Color Then
            GetCellColorString = cell.Offset(0, -1).Value
            Exit Function
        End If
    Next
    
    For Each cell In Worksheets("Settings").Range("G2", "G6")
        If cell.Interior.Color = Color Then
            GetCellColorString = "light " & cell.Offset(0, -2).Value
            Exit Function
        End If
    Next
End Function

Public Function GetColor(Color As String)
    For Each cell In Worksheets("Settings").Range("E2", "E20")
        If LCase(cell.Value) = LCase(Color) Then
            GetColor = GetCellColor(cell.Offset(0, 1))
            Exit Function
        End If
    Next

    GetColor = GetColor("default")
End Function


'TODO fix this
Public Function SetColor(Color As String)
    Application.Caller.Interior.Color = GetColor(Color)
End Function



Public Function GetColorLight(Color As String)
    For Each cell In Worksheets("Settings").Range("E2", "E20")
        If LCase(cell.Value) = LCase(Color) Then
            GetColorLight = GetCellColor(cell.Offset(0, 2))
            Exit Function
        End If
    Next

    GetColorLight = GetColorLight("default")
End Function



Public Function GetColorList()
    'Get number of colors
    'Dim ws As Worksheet
    Dim LastRowIndex As Integer
    Dim numColors As Integer

    'ws = Worksheets("Settings")
    LastRowIndex = Worksheets("Settings").Cells(Worksheets("Settings").Rows.count, "E").End(xlUp).row
    numColors = LastRowIndex - 1

    'Create empty array
    Dim ColorList() As String
    ReDim ColorList(numColors)

    Dim LastCell As String
    LastCell = "E" & CStr(LastRowIndex)

    Dim ColorListRange As Range
    Set ColorListRange = Worksheets("Settings").Range("E2", LastCell)

    'fill array
    Dim cell As Range

    For i = 1 To numColors
       Set cell = ColorListRange.Item(i)
       ColorList(i) = cell.Value
    Next

    GetColorList = ColorList
End Function


