Attribute VB_Name = "Testing"

Public Sub testEnum2()
    Dim test As New Collection
    
    Dim var As New GenVariable
    var.init "varTest1", varDate, strVarTypeOverride:="Banana"
    
    
    'var.VarTypeString = "Boolean"
    'MsgBox var.VarType & " " & var.VarTypeString
'    MsgBox var.CodeString
    
    Dim fTest As New GenFunction
    fTest.init "testFunc", fncFunction, varBoolean, strReturnTypeOverride:="Banana2"
    fTest.AddParameter "testParam", varByte, pParamType:=prmByVal, strVarTypeOverride:="manyBANANA"
    
    MsgBox fTest.CodeString
    
    'MsgBox test.Item
    
    'MsgBox test("test").Name
    
End Sub

Sub test5()
    Dim strProperty As String
    Dim vType As enmVarType
    Dim strVarName As String
    Dim aLevel As enmAccessLevel
    Dim strInheritedFrom As String
    Dim strVarTypeOverride As String

    strProperty = "PropertyTest2"
    vType = 0
    strVarName = "mPropTest"
    aLevel = acsPublic
    strVarTypeOverride = "lels"
    
    Dim fGet As New GenFunction
    Dim fSet As New GenFunction
    
'    Dim strVarType As String
'    strVarType = CodeEnums.GetVarTypeString(vType)
    
    If strVarName = "" Then
        'strVarName = "m" & GetPrefix(strVarType) & strProperty
        strVarName = "m" & strProperty
    End If
    
    'Me.AddVariable strVarName, vType, strVarTypeOverride:=strVarTypeOverride
    
    fGet.init strProperty, fncPropertyGet, vType, aLevel, strVarTypeOverride
    fGet.AddLine strProperty & " = " & strVarName
    
    fSet.init strProperty, fncPropertySet, varBlank, aLevel, strVarTypeOverride
    fSet.AddParameter "val", vType, pParamType:=prmByVal
    fSet.AddLine strVarName & " = val"
    
    MsgBox fGet.CodeString
    
'    Me.AddFunctionAsObj fGet, "PropertyGet_" & fGet.Name
'    Me.AddFunctionAsObj fSet, "PropertySet_" & fSet.Name
End Sub

Public Sub testArray()
'    Dim fTest As GenFunction
'    Set fTest = New GenFunction
'
'    fTest.init "TestFunc", fncFunction, varDate
'
'    fTest.AddParameter "param1", varString, True
'    fTest.AddParameter "param2", varBoolean, False
'    fTest.AddLine "Dim test As Integer"
'    fTest.AddLine "test = ""banananas Are greeeeaaaaat"""
    
    
'    Dim var As GenVariable
'    Set var = fTest.GetParameter("param2")
    
    
    Dim mTest As New GenModule
    mTest.init "TestModule"
    Dim var As New GenVariable
    var.init "testVar", varBoolean
'    mTest.Variables("your mom") = var
    mTest.AddVariable "testVar2", varByte
    mTest.AddSub "TestSub"
    mTest.Functions("TestSub").Name = "nope"
    'MsgBox mTest.Functions("nope").Name
    'mTest.Variables("testVar2").AccessLevel = acsPublic
    'MsgBox mTest.Variables("testVar2").AccessLevel
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
    'arrTableArray = SelectFromTableArrayTest(strTable, strFilterColumn, strFilterValue)
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
