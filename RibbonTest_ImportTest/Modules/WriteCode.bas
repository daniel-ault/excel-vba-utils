Attribute VB_Name = "WriteCode"
Sub WriteCode()
    Dim strPath, strFile As String
    strPath = ActiveWorkbook.path & "\Generated Code\"
    strFile = "test.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(strPath & strFile)
    
    'oFile.WriteLine "testing whoo"
    'oFile.WriteLine GenerateProperty("testProperty", "String")
    oFile.WriteLine GenerateControlCode("Checkbox")
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Function AddModule(ByVal strName As String, _
                   ByVal ComponentType As vbext_ComponentType, _
                   Optional ByVal blnDeleteIfExists As Boolean = True) As CodeModule
    Dim blnExists As Boolean
    Dim vbProj As VBIDE.VBProject
    Dim vbComps As VBIDE.vbComponents
    Dim vbComp As New VBIDE.vbComponent

    Set vbProj = ActiveWorkbook.VBProject
    Set vbComps = vbProj.vbComponents
    'Set vbComp = New vbComponent
    
    'Dim col As Collection
    'Set col = Convert.ChangeType(vbComps, Collection)
    blnExists = CollectionContains(vbComps, strName)
    If blnExists Then
        Set vbComp = vbComps(strName)
        
        If blnDeleteIfExists Then vbComps.Remove vbComp
    End If
    
    If Not blnExists Or blnDeleteIfExists Then
        Set vbComp = vbComps.Add(ComponentType)
        vbComp.Name = strName
    End If
    
    Set AddModule = vbComp.CodeModule
End Function

Sub AddCodeToModule(ByVal strName As String, ByRef strCode As String)
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.vbComponent
    Dim xMod As VBIDE.CodeModule
    
    With ActiveWorkbook
        Set xPro = .VBProject
        Set xCom = xPro.vbComponents(strName)
        Set xMod = xCom.CodeModule
        
        With xMod
            xMod.AddFromString strCode
        End With
    End With
End Sub

Sub GenerateAllModules()
    'Generate Control Module(s)
    
    'Generate Enum Module
    Dim strEnumModuleName As String
    strEnumModuleName = "CodeEnums"
    AddModule strEnumModuleName, vbext_ct_StdModule
    AddCodeToModule strEnumModuleName, GenerateAllEnums("tblEnumCode")
End Sub

Sub GenerateAllClasses()
    Dim arrClasses As Variant
    arrClasses = SelectTable("tblCodeClass")
    
    Dim i As Integer
    For i = LBound(arrClasses, 1) To UBound(arrClasses, 1)
        Dim strClassType, strClassName As String
        strClassType = arrClasses(i, 1)
        strClassName = arrClasses(i, 2)
        
        AddModule strClassName, vbext_ct_ClassModule
        AddCodeToModule strClassName, GenerateClassCodeAlt(strClassType)
    Next i
End Sub

Sub GenerateControlModule()
    Dim strControl As String
    strControl = "ToggleButton"
    
    AddModule strControl, vbext_ct_ClassModule
    AddCodeToModule strControl, GenerateControlCode(strControl)
End Sub

Function GenerateControlCode(strControl As String) As String
    'Get list of member variables
    '   retrieve from callback properties
    Dim strClass As String
    strClass = ""

    Dim arrAttributes As Variant
    arrAttributes = SelectFromTable("tblControlToAttribute", "strControl", strControl)
    
    'Generate member variables from attributes
    'TODO add callback properties that aren't attributes, like getPressed
    Dim i As Integer
    Dim strParam, strVar, strVarType, strAttribute As String
    Dim arrVarType As Variant
    Dim blnByRef As Boolean
    For i = LBound(arrAttributes, 1) To UBound(arrAttributes, 1)
        strAttribute = arrAttributes(i, 2)
        strVar = FirstUpper(strAttribute)
        
        arrVarType = SelectFromTable("tblAttributes", "strAttribute", strAttribute)
        strVarType = arrVarType(1, 2)
        
        strClass = strClass & "Private m" & GetPrefix(strVarType) & strVar & " As " & strVarType & vbCrLf
    Next i
    
    strClass = strClass & vbCrLf & vbCrLf
    
    For i = LBound(arrAttributes, 1) To UBound(arrAttributes, 1)
        strAttribute = arrAttributes(i, 2)
        strVar = UCase(Left(strAttribute, 1)) & Right(strAttribute, Len(strAttribute) - 1)
        
        arrVarType = SelectFromTable("tblAttributes", "strAttribute", strAttribute)
        strVarType = arrVarType(1, 2)

        strClass = strClass & GenerateProperty(strAttribute, strVarType)
        strClass = strClass & vbCrLf
    Next i
    
    GenerateControlCode = strClass
    
    'generate xml?
End Function

Private Function GenerateClassCode(ByVal strClass As String, _
                                   Optional ByVal strTable As String = "tblCodeClassProperties") As String
    Dim strCode As String
    Dim arrAttributes As Variant
    
    arrProperties = SelectFromTable(strTable, "strClass", strClass)
    
    For i = LBound(arrProperties, 1) To UBound(arrProperties, 1)
        Dim strProperty, strVarName, strPrefix As String
        Dim arrVarType As Variant
        
        strProperty = arrProperties(i, 2)
        strVarName = FirstUpper(arrProperties(i, 3))
        
        If strVarName = "" Then strVarName = arrProperties(i, 2)
        
        If arrProperties(i, 2) = "" Then
            strVarType = arrProperties(i, 4)
            strPrefix = GetPrefix(strVarType)
        Else
            strVarType = strProperty
        End If
        
        strCode = strCode & "Private m" & strPrefix & strVarName & " As " & strVarType & vbCrLf
    Next i
    
    strCode = strCode & vbCrLf & vbCrLf
    
    For i = LBound(arrProperties, 1) To UBound(arrProperties, 1)
        strProperty = arrProperties(i, 2)
        strVarName = FirstUpper(arrProperties(i, 3))
        If strVarName = "" Then strVarName = arrProperties(i, 2)
        
        If arrProperties(i, 2) = "" Then
            strVarType = arrProperties(i, 4)
            strPrefix = GetPrefix(strVarType)
        Else
            strVarType = strProperty
        End If

        strCode = strCode & GenerateProperty(strVarName, strVarType)
        strCode = strCode & vbCrLf
    Next i
    
    GenerateClassCode = strCode
    
End Function

Private Function GenerateClassCodeAlt(ByVal strClass As String, _
                                      Optional ByVal strTable As String = "tblCodeClassProperties") As String
    'Dim strCode As String
    Dim arrAttributes As Variant
    
    Dim ClassModule As New GenModule
    ClassModule.init strClass, modClass
        
    arrProperties = SelectFromTable(strTable, "strClass", strClass)
    
    For i = LBound(arrProperties, 1) To UBound(arrProperties, 1)
        Dim strProperty, strVarName, strPrefix, strVarNameFull As String
        'Dim VarType As enmVarType
        Dim arrVarType As Variant
        
        strProperty = arrProperties(i, 2)
        strVarName = FirstUpper(arrProperties(i, 3))
        
        If strVarName = "" Then strVarName = arrProperties(i, 2)
        
        If arrProperties(i, 2) = "" Then
            strVarType = arrProperties(i, 4)
            'VarType = CodeEnums.GetVarTypeEnum(arrProperties(i, 4))
            'strPrefix = GetPrefix(strVarType)
        Else
            strVarType = strProperty
            'VarType = CodeEnums.GetVarTypeEnum(arrProperties(i, 4))
        End If
        
        strVarNameFull = "m" & strPrefix & strVarName
        
        'ClassModule.AddVariable strVarNameFull, varType, acsPrivate
        ClassModule.AddProperty strVarName, 0, strVarNameFull, strVarTypeOverride:=strVarType
        'strCode = strCode & "Private m" & strPrefix & strVarName & " As " & strVarType & vbCrLf
    Next i
    
    GenerateClassCodeAlt = ClassModule.CodeGen
    
End Function

Private Function GenerateAllEnums(Optional ByVal strEnumTable As String = "tblEnum") As String
    'generate enums from tables
    Dim strCode As String
    strCode = ""
    
    Dim arrValueTypes As Variant
    arrValueTypes = SelectColumn("tblValueType", _
                                 "strValueType", _
                                 strFilterColumn:="strDataType", _
                                 vntFilterValue:="Enum")
    'arrValueTypes = SelectFromTable("tblValueType", "strDataType", "Enum")
    
    Dim colEnums As Collection
    Dim arrEnums, arrEnumStrings, vntEnum As Variant
    arrEnums = Application.Range(strEnumTable).value
    
    'Column 1 is declarations, 2 is definitions, 3 is functions
    Dim arrArrayCode As Variant
    ReDim arrArrayCode(LBound(arrEnums, 1) To UBound(arrEnums, 1), 1 To 3)
    
    'Generate Enum types
    For i = LBound(arrEnums, 1) To UBound(arrEnums, 1)
        Dim strName, strTable, strColumn, strNamePrefix, strElementPrefix As String
        Dim strArrayName, strVarName As String
        strName = arrEnums(i, 1)
        strTable = arrEnums(i, 2)
        strColumn = arrEnums(i, 3)
        strNamePrefix = arrEnums(i, 4)
        strElementPrefix = arrEnums(i, 5)
        
        strArrayName = "arr" & arrEnums(i, 1) & "Strings"
        strVarName = strElementPrefix & strName
        
        AddLine strCode, _
                GenerateEnumFromTable(strName, _
                                      strTable, _
                                      strColumn, _
                                      strNamePrefix:=strNamePrefix, _
                                      strElementPrefix:=strElementPrefix, _
                                      intNewLines:=1)
        
        Dim arr As Variant
        arr = SelectColumn(strTable, strColumn)
        arrArrayCode(i, 1) = "Private " & strArrayName & "() As Variant"
        arrArrayCode(i, 2) = GenerateArrayHardcode(arr, strArrayName)
        arrArrayCode(i, 3) = GenerateArrayFunction(strArrayName, strVarName, strName)
    Next i
    AddLine strCode, intNewLines:=1
    
    For i = LBound(arrArrayCode, 1) To UBound(arrArrayCode, 1)
        AddLine strCode, arrArrayCode(i, 1)
    Next i
    AddLine strCode, "Private blnStringArraysSet As Boolean", intNewLines:=2
    
    AddLine strCode, "Private Sub SetStringArrays()"
    For i = LBound(arrArrayCode, 1) To UBound(arrArrayCode, 1)
        AddLine strCode, arrArrayCode(i, 2), intTabCount:=1
    Next i
    AddLine strCode, "blnArraysSet = True", intTabCount:=1
    AddLine strCode, "End Sub", intNewLines:=2
    
    For i = LBound(arrArrayCode, 1) To UBound(arrArrayCode, 1)
        AddLine strCode, arrArrayCode(i, 3), intNewLines:=1
    Next i

    GenerateAllEnums = strCode
End Function



Private Function GenerateEnumFromTable(ByVal strName As String, _
                                       ByVal strTable As String, _
                                       ByVal strColumn As String, _
                                       Optional ByVal strFilterColumn As String, _
                                       Optional ByVal vntFilterValue As Variant, _
                                       Optional ByVal strNamePrefix As String = "enm", _
                                       Optional ByVal strElementPrefix As String = "enm", _
                                       Optional ByVal intTabCount As Integer = 0, _
                                       Optional ByVal intNewLines As Integer = 1)
    GenerateEnumFromTable = GenerateEnum(strName, _
                                         SelectColumn(strTable, strColumn, strFilterColumn, vntFilterValue), _
                                         strNamePrefix:=strNamePrefix, _
                                         strElementPrefix:=strElementPrefix, _
                                         intTabCount:=intTabCount, _
                                         intNewLines:=intNewLines)
End Function

Private Function GenerateEnum(ByVal strName As String, _
                              arr As Variant, _
                              Optional ByVal strNamePrefix As String = "enm", _
                              Optional ByVal strElementPrefix As String = "enm", _
                              Optional ByVal intTabCount As Integer = 0, _
                              Optional ByVal intNewLines As Integer = 1)
                              
    Dim strCode As String
    strCode = ""
    
    AddLine strCode, _
            "Enum " & strNamePrefix & FirstUpper(strName), _
            intTabCount:=intTabCount
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> "" Then
            AddLine strCode, _
                    strElementPrefix & FirstUpper(arr(i)), _
                    intTabCount:=intTabCount + 1
        End If
    Next i
    
    AddLine strCode, "End Enum", intTabCount:=intTabCount, intNewLines:=intNewLines
    
    GenerateEnum = strCode
End Function

Private Function GenerateArrayHardcode(arr As Variant, _
                                       ByVal strName As String, _
                                       Optional ByVal intTabCount As Integer = 0, _
                                       Optional ByVal intNewLines As Integer = 1) As String
    Dim strCode As String
    
    strCode = strCode + strName & " = Array("
    
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim str As String
        str = """" & arr(i) & """"
        If i <> UBound(arr, 1) Then str = str + ", "
        
        strCode = strCode + str
    Next i
    
    strCode = strCode + ")"
    
    'AddLine strCode, ""
    
    GenerateArrayHardcode = strCode
End Function


Private Function GenerateArrayFunction(ByVal strArrayName As String, _
                                       ByVal strVarName As String, _
                                       ByVal strVarType As String, _
                                       Optional ByVal strFuncType As String = "String", _
                                       Optional ByVal strArrayType As String = "String", _
                                       Optional ByVal intTabCount As Integer = 0, _
                                       Optional ByVal intNewLines As Integer = 1) As String
    Dim strCode As String
    Dim strName As String
    strName = "Get" & strVarType & strArrayType
    
    AddLine strCode, "Public Function " & strName & "(" & strVarName & " As " & strVarType & ") As " & strFuncType, intTabCount
    AddLine strCode, "If Not bln" & strArrayType & "ArraysSet Then Set" & strArrayType & "Arrays", intTabCount + 1
    AddLine strCode, intTabCount:=intTabCount + 1
    AddLine strCode, strName & " = " & strArrayName & "(" & strVarName & ")", intTabCount + 1
    AddLine strCode, "End Function", intTabCount, intNewLines

    GenerateArrayFunction = strCode
End Function

Private Function GenerateAssocArrayFunction(ByVal strFuncName As String, _
                                            ByVal strVarName1 As String, _
                                            ByVal strVarType1 As String, _
                                            ByVal strVarName2 As String, _
                                            ByVal strVarType2 As String, _
                                            ByVal strArrayName As String, _
                                            Optional ByVal intTabCount As Integer = 0, _
                                            Optional ByVal intNewLines As Integer = 1) As String
    Dim strCode As String
    
    AddLine strCode, "Public Function " & strFuncName & "(" & strVarName1 & " As " & strVarType1 & ", " & strVarName2 & " As " & strVarType2 & ") As Boolean", intTabCount
    AddLine strCode, "If Not blnAssocArraysSet Then SetAssociationArrays", intTabCount + 1
    AddLine strCode
    AddLine strCode, strFuncName & " = " & strArrayName & "(" & strVarName2 & ", " & strVarName1 & ")", intTabCount + 1
    AddLine strCode, "End Function", intTabCount, intNewLines
End Function



Private Function GenerateProperty(ByVal strProperty As String, _
                                  ByVal strVarType As String, _
                                  Optional ByVal strInheritedFrom As String = "")
    Dim str As String
    str = ""
    
    'TODO change strProperty if strInheritedFrom has a value
    '   e.g. instead of Color, it would be XmlObject_Color
    
    strProperty = FirstUpper(strProperty)
    
    Dim strVarName As String
    strVarName = "m" & GetPrefix(strVarType) & strProperty
    
    str = str & "Public Property Get " & strProperty & "() As " & strVarType & vbCrLf
    str = str & vbTab & strProperty & " = " & strVarName & vbCrLf
    str = str & "End Property" & vbCrLf
    
    'str = str & vbCrLf
    
    str = str & "Public Property Let " & strProperty & "(ByVal val As " & strVarType & ")" & vbCrLf
    str = str & vbTab & strVarName & " = val" & vbCrLf
    str = str & "End Property" & vbCrLf
    
    GenerateProperty = str
End Function

Public Function FirstUpper(ByVal strString As String) As String
    If strString = "" Then Exit Function
    FirstUpper = UCase(Left(strString, 1)) & Right(strString, Len(strString) - 1)
End Function

Public Function CollectionContains(col As VBIDE.vbComponents, key As Variant) As Boolean
    Dim obj As Variant
On Error GoTo err
    CollectionContains = True
    Set obj = col(key)
    Exit Function
err:
    CollectionContains = False
End Function

Private Sub AddLine(ByRef strCode As String, _
                    Optional ByVal strLine As String = "", _
                    Optional ByVal intTabCount As Integer = 0, _
                    Optional ByVal intNewLines As Integer = 1)
    
    If intTabCount > 0 Then
        Dim i As Integer
        For i = 1 To intTabCount
            strCode = strCode + vbTab
        Next i
    End If
    
    strCode = strCode + strLine
    
    If intNewLines > 0 Then
        For i = 1 To intNewLines
            strCode = strCode + vbCrLf
        Next i
    End If
End Sub
