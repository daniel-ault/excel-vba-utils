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

Sub GenerateAllModules()
    'Generate Control Module(s)
    
    'Generate Enum Module
    Dim strEnumModuleName As String
    strEnumModuleName = "Enums"
    AddModule strEnumModuleName, vbext_ct_StdModule
    AddCodeToModule strEnumModuleName, GenerateAllEnums
End Sub

Sub GenerateControlModule()
    Dim strControl As String
    strControl = "Checkbox"
    
    AddModule strControl, vbext_ct_ClassModule
    AddCodeToModule strControl, GenerateControlCode(strControl)
End Sub

Sub AddModule(ByVal strName As String, _
              ByVal ComponentType As vbext_ComponentType, _
              Optional ByVal blnDeleteIfExists As Boolean = True)
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
        Set vbComp = vbComps.Add(vbext_ct_StdModule)
        vbComp.Name = strName
    End If
End Sub

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


Function GenerateControlCode(strControl As String) As String
    'Get list of member variables
    '   retrieve from callback properties
    Dim strClass As String
    strClass = ""

    Dim arrAttributes As Variant
    arrAttributes = SelectFromTable("tblControlToAttribute", "strControl", strControl)
    
    'Generate member variables from callback properties
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

Private Function GenerateAllEnums() As String
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
    Dim arrEnums, vntEnum As Variant
    arrEnums = Application.Range("tblEnum")
    
    For i = LBound(arrEnums, 1) To UBound(arrEnums, 1)
        
    Next i
    
    For Each vntEnum In colEnums
        Dim strName, strTable, strColumn, strNamePrefix, strElementPrefix As String
'        strName = colEnums(i, 1)
'        strTable = colEnums(i, 2)
'        strColumn = colEnums(i, 3)
'        strNamePrefix = colEnums(i, 4)
'        strElementPrefix = colEnums(i, 5)
        
        AddLine strCode, _
                GenerateEnumFromTable(strName, _
                                      strTable, _
                                      strColumn, _
                                      strNamePrefix:=strNamePrefix, _
                                      strElementPrefix:=strElementPrefix, _
                                      intNewLines:=1)
    Next vntEnum
    
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
                    ByVal strLine As String, _
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
