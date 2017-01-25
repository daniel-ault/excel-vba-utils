Attribute VB_Name = "CodeEnums"
Enum enmAccessLevel
    acsBlank
    acsPublic
    acsPrivate
    acsDim
End Enum

Enum enmVarType
    varBlank
    varVariant
    varByte
    varBoolean
    varInteger
    varLong
    varSingle
    varDouble
    varCurrency
    varDecimal
    varDate
    varObject
    VarString
    varCollection
    varOverride
End Enum

Enum enmParamType
    prmBlank
    prmByRef
    prmByVal
End Enum

Enum enmModuleType
    modClass
    modModule
End Enum

Enum enmFunctionType
    fncSub
    fncFunction
    fncPropertyGet
    fncPropertyLet
    fncPropertySet
End Enum

'TODO add vartype prefixes?

Private arrAccessLevelStrings() As Variant
Private arrVarTypeStrings() As Variant
Private arrParamTypeStrings() As Variant
Private arrModuleTypeStrings() As Variant
Private arrFunctionTypeStrings() As Variant
Private blnStringArraysSet As Boolean

Private Sub SetStringArrays()
    arrAccessLevelStrings = Array("", "Public", "Private", "Dim")
    arrVarTypeStrings = Array("", "Variant", "Byte", "Boolean", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Object", "String", "Collection", "Override")
    arrParamTypeStrings = Array("", "ByRef", "ByVal")
    arrModuleTypeStrings = Array("Class", "Module")
    arrFunctionTypeStrings = Array("Sub", "Function", "Property Get", "Property Let", "Property Set")
    blnArraysSet = True
End Sub

Public Function GetAccessLevelString(acsAccessLevel As enmAccessLevel) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetAccessLevelString = arrAccessLevelStrings(acsAccessLevel)
End Function

Public Function GetVarTypeString(varVarType As enmVarType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetVarTypeString = arrVarTypeStrings(varVarType)
End Function

Public Function GetVarTypeEnum(ByVal strVarType As String) As enmVarType
    If Not blnStringArraysSet Then SetStringArrays
    
    Dim i As Integer
    For i = LBound(arrVarTypeStrings) To UBound(arrVarTypeStrings)
        If arrVarTypeStrings(i) = strVarType Then
            GetVarTypeEnum = i
            Exit Function
        End If
    Next i
    
    GetVarTypeEnum = -1
End Function

Public Function GetParamTypeString(prmParamType As enmParamType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetParamTypeString = arrParamTypeStrings(prmParamType)
End Function

Public Function GetModuleTypeString(modModuleType As enmModuleType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetModuleTypeString = arrModuleTypeStrings(modModuleType)
End Function

Public Function GetFunctionTypeString(fncFunctionType As enmFunctionType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetFunctionTypeString = arrFunctionTypeStrings(fncFunctionType)
End Function

Public Function GetOptionalString(blnIsOptional As Boolean)
    If blnIsOptional Then
        GetOptionalString = "Optional"
    Else
        GetOptionalString = ""
    End If
End Function

