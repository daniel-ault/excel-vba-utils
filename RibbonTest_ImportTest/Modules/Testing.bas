Attribute VB_Name = "Testing"

Public Sub testEnum2()
    Dim test As Collection
    Set test = New Collection
    
    test.Add "Name"
    test.Add "item"
    test.Add "whoa"
    
    Dim vnt As Variant
    For Each vnt In test
        MsgBox vnt
    Next vnt
End Sub

Public Sub TestSelects()
    Dim strTable, strColumn, strFilterColumn, strFilterValue As String
    Dim intColumn, intFilterColumn, intTableColumn As Integer
    Dim tblControlToAttribute As Range
    Dim arrControlToAttribute As Variant
    Dim sw As StopWatch: Set sw = New StopWatch
    
    Dim lngTimerTable, lngTimerTableArray, lngTimerArray As Long
    Dim arrTable, arrTableArray, arrArray As Variant
    
    strTable = "tblControlToAttribute"
    strColumn = "strAttribute"
    strFilterColumn = "strControl"
    strFilterValue = "Button"
    
    Set tblControlToAttribute = Application.Range(strTable)
    arrControlToAttribute = tblControlToAttribute.value
    
    intTableColumn = tblControlToAttribute.Column
    intColumn = Application.Range(strTable & "[" & strColumn & "]").Column
    intColumn = intColumn - intTableColumn + 1
    intFilterColumn = Application.Range(strTable & "[" & strFilterColumn & "]").Column
    intFilterColumn = intFilterColumn - intTableColumn + 1
    
    sw.StartTimer
    arrTable = SelectFromTable(strTable, strFilterColumn, strFilterValue)
    lngTimerTable = sw.EndTimer
    
    sw.StartTimer
'    arrTableArray = SelectFromTableArrayTest(strTable, strFilterColumn, strFilterValue)
    lngTimerTableArray = sw.EndTimer
    
    sw.StartTimer
    arrArray = SelectFromArray(arrControlToAttribute, intFilterColumn, strFilterValue)
    lngTimerArray = sw.EndTimer
End Sub


Public Sub TestBinary()
    Dim b1, b2, b3 As Long
    b1 = 15     '0000 1111
    b2 = 120    '0111 1000
    
    'should result in 0000 1000, or 8
    b3 = b1 And b2
End Sub

Public Sub AddAttributes()
    Dim strControl, strList As String
    strControl = InputBox("Enter Control name")
    strList = InputBox("Enter list of attributes")
    
    strList = Replace(strList, """", "")
    strList = Replace(strList, " ", "")
    strList = Replace(strList, vbCrLf, "")
    
    Dim arr As Variant
    
    arr = Split(strList, ",")
    
    If Not ExistsInTable("tblControl", "strControl", strControl) Then
        AddControl (strControl)
    End If
    
    Dim str As Variant
    For Each str In arr
        If Not str = "" Then
            AddControlToAttribute strControl, str
        End If
    Next str
    
    'MsgBox UBound(arr)
End Sub



Sub AddControl(strControl As String)
    Dim tbl As Range
    Set tbl = Application.Range("tblControl")
    
    tbl.End(xlDown).Offset(1).value = strControl
End Sub

Sub AddControlToAttribute(ByVal strControl As String, ByVal strAttribute As String)
    Dim tbl As Range
    Set tbl = Application.Range("tblControlToAttribute")
    
    Dim rngEnd As Range
    Set rngEnd = tbl.End(xlDown).Offset(1)
    
    rngEnd.value = strControl
    rngEnd.Offset(0, 1).value = strAttribute
End Sub

Sub AddControlToCallback(ByVal strControl As String, ByVal strCallback As String)
    Dim tbl As Range
    Set tbl = Application.Range("tblControlToCallback")
    
    Dim rngEnd As Range
    Set rngEnd = tbl.End(xlDown).Offset(1)
    
    rngEnd.value = strControl
    rngEnd.Offset(0, 1).value = strCallback
End Sub

Sub test2()
    MsgBox ExistsInTable("tblControl", "strControl", "Checkbox")
End Sub
