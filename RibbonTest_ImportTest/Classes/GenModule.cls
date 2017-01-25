VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mModuleType As enmModuleType
Private mAccessLevel As enmAccessLevel
Private mstrName As String
Private mVariables As Collection
Private mFunctions As Collection

Private Sub Class_Initialize()
    mModuleType = 0
    mAccessLevel = 0
    Set mVariables = New Collection
    Set mFunctions = New Collection
End Sub


Public Sub init(ByVal strName As String, _
                Optional ByVal mModuleType As enmModuleType = modModule, _
                Optional ByVal aAccessLevel As enmAccessLevel = acsPrivate)
    Me.Name = strName
    Me.ModuleType = mModuleType
    Me.AccessLevel = aAccessLevel
End Sub

Public Sub AddVariable(ByVal strName As String, _
                       ByVal vType As enmVarType, _
                       Optional ByVal aLevel As enmAccessLevel = acsPrivate, _
                       Optional ByVal strVarTypeOverride As String)
    Dim var As GenVariable
    Set var = New GenVariable
    var.init strName, vType, aLevel, strVarTypeOverride:=strVarTypeOverride
    
    mVariables.Add var, var.Name
End Sub

Public Sub AddVariableAsObj(ByRef var As GenVariable)
    mVariables.Add var, var.Name
End Sub

Public Sub RemoveVariable(ByVal strName As String)
    mVariables.Remove strName
End Sub

Public Function GetVariable(ByVal strName As String) As GenVariable
    Set GetVariables = mVariables.Item(strName)
End Function


Public Sub AddSub(ByVal strName As String, _
                  Optional ByVal aLevel As enmAccessLevel = acsPrivate)
    Dim fSub As New GenFunction
    fSub.init strName, fncSub, level:=aLevel
    mFunctions.Add fSub, fSub.Name
    
    'MsgBox mFunctions(fSub.Name).Name
End Sub

Public Sub AddFunction(ByVal strName As String, _
                       ByVal fType As enmFunctionType, _
                       Optional ByVal vReturnType As enmVarType, _
                       Optional ByVal level As enmAccessLevel = acsBlank, _
                       Optional ByVal strVarTypeOverride As String)
    Dim fFunc As New GenFunction
    fFunc.init strName, fType, vReturnType, level, strVarTypeOverride
    mFunctions.Add fFunc, fFunc.Name
End Sub

Public Function AddFunctionAsObj(ByRef fFunc As GenFunction, _
                                 Optional ByVal key As Variant)
    If IsMissing(key) Then key = fFunc.Name
    mFunctions.Add fFunc, key
End Function

Public Sub RemoveFunction(ByVal strName As String)
    mFunctions.Remove strName
End Sub


Public Function CodeGen() As String
    Dim strCode As String
    
    Dim var As GenVariable
    For Each var In mVariables
        strCode = strCode & var.CodeString & vbCrLf
    Next var
    
    strCode = strCode & vbCrLf & vbCrLf
    
    Dim func As GenFunction
    Dim strNewLines As String
    For Each func In mFunctions
        If func.FunctionType = fncPropertyGet Then
            strNewLines = vbCrLf
        ElseIf func.FunctionType = fncPropertyLet Or func.FunctionType = fncPropertySet Then
            strNewLines = vbCrLf & vbCrLf
        Else
            strNewLines = vbCrLf & vbCrLf & vbCrLf
        End If
        strCode = strCode & func.CodeString & strNewLines
    Next func

    CodeGen = strCode
End Function

Public Sub AddProperty(ByVal strProperty As String, _
                       ByVal vType As enmVarType, _
                       Optional ByVal strVarName As String, _
                       Optional ByVal aLevel As enmAccessLevel = acsPublic, _
                       Optional ByVal strInheritedFrom As String, _
                       Optional ByVal strVarTypeOverride As String)
    Dim fGet As New GenFunction
    Dim fLet As New GenFunction
    
'    Dim strVarType As String
'    strVarType = CodeEnums.GetVarTypeString(vType)
    
    If strVarName = "" Then
        'strVarName = "m" & GetPrefix(strVarType) & strProperty
        strVarName = "m" & strProperty
    End If
    
    Me.AddVariable strVarName, vType, strVarTypeOverride:=strVarTypeOverride
    
    fGet.init strProperty, fncPropertyGet, vType, aLevel, strVarTypeOverride
    fGet.AddLine strProperty & " = " & strVarName
    
    fLet.init strProperty, fncPropertyLet, varBlank, aLevel, strVarTypeOverride
    fLet.AddParameter "val", vType, pParamType:=prmByVal, strVarTypeOverride:=strVarTypeOverride
    fLet.AddLine strVarName & " = val"
    
    Me.AddFunctionAsObj fGet, "PropertyGet_" & fGet.Name
    Me.AddFunctionAsObj fLet, "PropertyLet_" & fLet.Name
End Sub

Public Sub RemoveProperty(ByVal strProperty As String, _
                          Optional ByVal blnDeleteVariable As Boolean = True, _
                          Optional ByVal strVarName As String)
    Me.RemoveFunction "PropertyGet_" & strProperty
    Me.RemoveFunction "PropertyLet_" & strProperty
    
    If blnDeleteVariable Then
        If strVarName = "" Then strVarName = "m" & strVarName
        Me.RemoveVariable strVarName
    End If
    
End Sub

'TODO currently, if you reset the name via Functions(strName).Name, the key stays the same
'     possibly create custom Collection class for variables and Functions?
Public Property Get Functions(ByVal strName) As GenFunction
    Set Functions = mFunctions(strName)
End Property


Public Property Get ModuleType() As enmModuleType
    ModuleType = mModuleType
End Property
Public Property Let ModuleType(ByVal val As enmModuleType)
    mModuleType = val
End Property

Public Property Get AccessLevel() As enmAccessLevel
    AccessLevel = mAccessLevel
End Property
Public Property Let AccessLevel(ByVal val As enmAccessLevel)
    mAccessLevel = val
End Property

Public Property Get Name() As String
    Name = mstrName
End Property
Public Property Let Name(ByVal val As String)
    mstrName = val
End Property


'Public Property Get Variables(ByVal strName As String) As GenVariable
'    Variables = mVariables(strName)
'End Property
'Public Property Set Variables(ByVal strName As String, _
'                              ByRef variable As GenVariable)
'    mVariables.Item(strName) = variable
'End Property
'
'Public Property Get Functions() As enmFunctionType
'    Functions = mFunctions
'End Property
'Public Property Let Functions(ByVal val As enmFunctionType)
'    mFunctions = val
'End Property


