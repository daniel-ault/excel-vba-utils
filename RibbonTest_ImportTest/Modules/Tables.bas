Attribute VB_Name = "Tables"
Sub LoadCustRibbon()

    Dim hFile As Long
    Dim path As String, fileName As String, ribbonXML As String, user As String
    
    hFile = FreeFile
    user = Environ("Username")
    path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
    fileName = "Excel.officeUI"
    
    ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
    ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='Reports' insertBeforeQ='mso:TabHome'>" & vbNewLine
    ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='Reports' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='runReport' label='PTO' " & vbNewLine
    ribbonXML = ribbonXML + "imageMso='AppointmentColor3'      onAction='GenReport'/>" & vbNewLine
    ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
    ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
    ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "</mso:customUI>"
    
    ribbonXML = Replace(ribbonXML, """", "")
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile

End Sub


Sub testFilter()
    'Application.Range("tblValueTypeToValue").AutoFilter Field:=1, Criteria1:="enmAlignment"
'    Application.Range("tblValueTypeToValue").AdvancedFilter _
'        xlFilterInPlace, _
'        Application.Range("tblValueType[strValueType]")
End Sub


Function AssociationExists(ByVal strTable As String, _
                           vntValue1 As Variant, _
                           vntValue2 As Variant, _
                           Optional ByVal strColumn1 As String = "", _
                           Optional ByVal strColumn2 As String = "") As Boolean
    Dim tbl As Range
    Dim arr As Variant
    Dim i, intColumn1, intColumn2 As Integer
    
    Set tbl = Application.Range(strTable)
    
    If strColumn1 <> "" Then
        intColumn1 = GetColumnInt(strTable, strColumn1)
    Else
        intColumn1 = 1
    End If
    
    If strColumn2 <> "" Then
        intColumn2 = GetColumnInt(strTable, strColumn2)
    Else
        intColumn2 = 2
    End If
    
    arr = tbl.value
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If (arr(i, intColumn1) = vntValue1) And (arr(i, intColumn2) = vntValue2) Then
            AssociationExists = True
            Exit Function
        End If
    Next i
    
    AssociationExists = False
End Function
                           
Function GetColumnInt(strTable As String, strColumn As String) As Integer
    Dim tbl As Range
    Dim intColumn As Integer
    
    tbl = Application.Range(strTable)
    
    intColumn = Application.Range(strTable & "[" & strColumn & "]").Column
    intColumn = intColumn - tbl.Column + 1
    
    GetColumnInt = intColumn
End Function


Function SelectColumn(ByVal strTable As String, _
                      ByVal strColumn As String, _
                      Optional strFilterColumn As String, _
                      Optional vntFilterValue As Variant)
    Dim arr As Variant
    Dim intCol As Integer
    
    'If Not vntFilterValue Is Missing Then
    If Not IsMissing(vntFilterValue) Then
        If strFilterColumn = "" Then strFilterColumn = strColumn
        
        arr = SelectFromTable(strTable, strFilterColumn, vntFilterValue)
        
        Dim tbl As Range
        Set tbl = Application.Range(strTable)
        intCol = Application.Range(strTable & "[" & strColumn & "]").Column
        intCol = intCol - tbl.Column + 1
    Else
        arr = Application.Range(strTable & "[" & strColumn & "]").value
        intCol = 1
    End If
    
    With Application
        SelectColumn = .Transpose(.Index(arr, 0, intCol))
    End With
End Function



Function SelectColumnFromArray(arr As Variant, _
                               ByVal intColumn As Integer, _
                               Optional ByVal intFilterColumn As Integer, _
                               Optional ByVal vntFilterValue As Variant)
    Dim arrFiltered As Variant
        
    If Not IsNull(vntFilterValue) Then
        If intFilterColumn = 0 Then intFilterColumn = intColumn
        
        arrFiltered = SelectFromArray(arr, intFilterColumn, vntFilterValue)
    End If
    
    With Application
        SelectColumnFromArray = .Transpose(.Index(arr, 0, intColumn))
    End With
End Function

'Function SelectFromTable(ByVal strTable As String, _
'                         ByVal strFilterColumn As String, _
'                         ByVal vntFilterValue As Variant) As Variant
'    Dim tbl As Range
'    Set tbl = Application.Range(strTable)
'    Set tbl = Worksheets(tbl.Worksheet.Name).Range(tbl.Address)
'
'    Dim strWorksheet As String
'    strWorksheet = tbl.Worksheet.Name
'
'    Dim intCol As Integer
'    intCol = Application.Range(strTable & "[" & strFilterColumn & "]").Column
'    intCol = intCol - tbl.Column + 1
'
'    'tbl.Worksheet.EnableAutoFilter = True
'    'tbl.Worksheet.ShowAllData
'    tbl.AutoFilter Field:=intCol
'    'tbl.Sort Key1:=tbl.Cells(2, intCol), Order1:=xlAscending, Header:=xlYes
'    tbl.Sort Key1:=tbl, Order1:=xlAscending, Header:=xlYes
'    tbl.AutoFilter Field:=intCol, Criteria1:=vntFilterValue
'
'    Dim tblFullTable As Range
'    Set tblFullTable = Application.Range(strTable & "[#ALL]").SpecialCells(xlCellTypeVisible)
'
'    If tblFullTable.Areas.count = 1 And tblFullTable.Rows.count = 1 Then
'        'SelectFromTable = Nothing
'        tbl.AutoFilter Field:=intCol
'        Exit Function
'    End If
'
'    Dim tblFiltered As Range
'    Set tblFiltered = tbl.SpecialCells(xlCellTypeVisible)
'
'    SelectFromTable = tblFiltered.value
'    tbl.AutoFilter Field:=intCol
'End Function

Function SelectTable(ByVal strTable As String) As Variant
    SelectTable = Application.Range(strTable).value
End Function

Function SelectFromTable(ByVal strTable As String, _
                         ByVal strFilterColumn As String, _
                         ByVal vntFilterValue As Variant) As Variant
    Dim tbl As Range
    Dim intCol As Integer
    
    Set tbl = Application.Range(strTable)
    intCol = Application.Range(strTable & "[" & strFilterColumn & "]").Column
    intCol = intCol - tbl.Column + 1
    
    SelectFromTable = SelectFromArray(tbl.value, intCol, vntFilterValue)
End Function

Function SelectFromArray(arr As Variant, _
                         ByVal intFilterColumn As Integer, _
                         ByVal vntFilterValue As Variant)
                         
    Dim arrFiltered As Variant
    Dim intCount, i, j As Integer: intCount = LBound(arr, 1) - 1
    ReDim arrFiltered(LBound(arr, 1) To UBound(arr, 1), _
                      LBound(arr, 2) To UBound(arr, 2))
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, intFilterColumn) = vntFilterValue Then
            intCount = intCount + 1
            For j = LBound(arr, 2) To UBound(arr, 2)
                arrFiltered(intCount, j) = arr(i, j)
            Next j
        End If
    Next i
    
    arrFiltered = RedimMulti(arrFiltered, intCount)
    
    SelectFromArray = arrFiltered
End Function


Sub TestGetCallbackParams()
    Dim arr As Variant
    arr = GetAllCallbackParams
    
    Dim j As Integer
    For j = LBound(arr, 1) To UBound(arr, 1)
        MsgBox arr(j, 1) & " " & arr(j, 2) & " " & arr(j, 3) & " " & arr(j, 4)
    Next j
End Sub

Function ExistsInTable(ByVal strTable As String, _
                       ByVal strFilterColumn As String, _
                       ByVal value As Variant)
    Dim tbl As Range
    Set tbl = Application.Range(strTable & "[" & strFilterColumn & "]")
    
    Dim rng As Range
    For Each rng In tbl.Cells
        If LCase(rng.value) = LCase(value) Then
            ExistsInTable = True
            Exit Function
        End If
    Next rng
    
    ExistsInTable = False
End Function

Function GetAllCallbackParams(Optional ByVal strControl As String = "Checkbox") As Variant
    Dim arrCallbacks, arrCallbackParams As Variant
    Dim numCallbacks, count As Integer
    
    arrCallbacks = SelectFromTable("tblControlToCallback", "strControl", strControl)
    numCallbacks = UBound(arrCallbacks) - LBound(arrCallbacks) + 1
    
    ReDim arrCallbackParams(1 To numCallbacks, 1 To 5)
    
    count = 0
    Dim i As Integer
    For i = LBound(arrCallbacks) To UBound(arrCallbacks)
        Dim arrParam As Variant
        Dim strCallback As String
        strCallback = arrCallbacks(i, 2)

        arrParam = SelectFromTable("tblCallbackParams", "strCallback", strCallback)

        'If Not IsArray(arrCallbackParams) Then GoTo ContinueFor
        If IsArray(arrParam) Then
            Dim j As Integer
            For j = LBound(arrParam) To UBound(arrParam)
                Dim strParam, strParamType As String
                Dim blnByRef As Boolean
                count = count + 1
                If count > numCallbacks Then
                    numCallbacks = numCallbacks + 10
                    arrCallbackParams = RedimMulti(arrCallbackParams, numCallbacks)
                End If
                arrCallbackParams(count, 1) = arrParam(j, 1)
                arrCallbackParams(count, 2) = arrParam(j, 2)
                arrCallbackParams(count, 3) = arrParam(j, 3)
                arrCallbackParams(count, 4) = arrParam(j, 4)
                
                'MsgBox strParam & " " & strParamType & " " & blnByRef
            Next j
        End If
        'MsgBox arrCallbacks(i, 2)
'ContinueFor:
    Next i
    
    GetAllCallbackParams = RedimMulti(arrCallbackParams, count)
End Function

Function RedimMulti(ByRef arr As Variant, ByVal nSize As Integer)
    Dim tmp As Variant
    ReDim tmp(1 To nSize, 1 To UBound(arr, 2))
    Dim i As Integer
    
    Dim n As Integer
    If UBound(arr, 1) < nSize Then
        n = UBound(arr, 1)
    Else
        n = nSize
    End If
    
    For i = LBound(arr) To n
        Dim j As Integer
        For j = LBound(arr, 2) To UBound(arr, 2)
            tmp(i, j) = arr(i, j)
        Next j
    Next i
    
    RedimMulti = tmp
End Function

Function GetPrefix(ByVal strVarType As String, _
                   Optional ByVal strTable As String = "tblVarType")
    Dim tbl As Range
    Set tbl = Application.Range(strTable)
    
    Dim r As Range
    For Each r In tbl.Rows
        If r.Cells(1).value = strVarType Then
            GetPrefix = r.Cells(2).value
            Exit Function
        End If
    Next r
    
    GetPrefix = ""
End Function


Function GetCallbackParams(strCallback As String) As Variant
    Dim tbl As Range
    Set tbl = Application.Range("tblCallbackParams")
    tbl.AutoFilter Field:=1, Criteria1:=strControl
    
    Dim row As Range
    For Each row In tbl.Rows
        If Not row.EntireRow.Hidden Then
            MsgBox row.Cells(ColumnIndex:=2)
        End If
    Next row
End Function

Sub ControlToCallback(strControl As String)
    Dim tbl As Range
    Set tbl = Application.Range("tblControlToCallback")
    tbl.AutoFilter Field:=1, Criteria1:=strControl
    
    Dim row As Range
    For Each row In tbl.Rows
        If Not row.EntireRow.Hidden Then
            MsgBox row.Cells(ColumnIndex:=2)
        End If
    Next row
    
    tbl.Worksheet.ShowAllData
End Sub


'Sub GetCallbackParams() 'strControl As String)
'    Dim strControl As String
'    strControl = "Checkbox"
'
'
'End Sub
