VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mFunctionType As enmFunctionType
Private mAccessLevel As enmAccessLevel
Private mVarType As enmVarType
Private mstrName As String

Private mParameters As Collection
Private mstrCode As String

Private mstrVarTypeOverride As String

'TODO possibly add comments to functions? simply placed above them

Private Sub Class_Initialize()
    mFunctionType = 0
    mAccessLevel = 0
    mVarType = 0
    Set mParameters = New Collection
End Sub

Public Sub init(ByVal strName As String, _
                ByVal fType As enmFunctionType, _
                Optional ByVal vReturnType As enmVarType, _
                Optional ByVal level As enmAccessLevel = acsBlank, _
                Optional ByVal strReturnTypeOverride As String)
    Me.Name = strName
    Me.FunctionType = fType
    If strReturnTypeOverride <> "" Then
        VarTypeString = strReturnTypeOverride
    Else
        Me.VarType = vReturnType
    End If
    Me.AccessLevel = level
End Sub

Public Sub AddParameter(ByVal strName As String, _
                        ByVal vType As enmVarType, _
                        Optional ByVal blnIsOptional As Boolean = False, _
                        Optional ByVal pParamType As enmParamType = prmBlank, _
                        Optional ByVal strVarTypeOverride As String)
    Dim param As New GenVariable
    param.init strName, vType, acsBlank, True, blnIsOptional, pParamType, strVarTypeOverride
    
    mParameters.Add param, strName
End Sub

Public Sub RemoveParameter(ByVal strName As String)
    mParameters.Remove strName
End Sub

Public Function GetParameter(ByVal strName As String) As GenVariable
    Set GetParameter = mParameters.Item(strName)
End Function

Public Sub AddLine(Optional ByVal strLine As String = "", _
                   Optional ByVal intTabCount As Integer = 1, _
                   Optional ByVal intNewLines As Integer = 1)
    
    If intTabCount > 0 Then
        Dim i As Integer
        For i = 1 To intTabCount
            mstrCode = mstrCode + vbTab
        Next i
    End If
    
    mstrCode = mstrCode + strLine
    
    If intNewLines > 0 Then
        For i = 1 To intNewLines
            mstrCode = mstrCode + vbCrLf
        Next i
    End If
End Sub

'TODO Public Sub DeleteLine ?
Public Function DeclarationString() As String
    Dim strDeclaration As String
    Dim param As GenVariable
    'strDeclaration = ""
    strDeclaration = strDeclaration & AccessLevelString
    strDeclaration = strDeclaration & FunctionTypeString
    strDeclaration = strDeclaration & Me.Name
    strDeclaration = strDeclaration & "("
    For Each param In mParameters
        strDeclaration = strDeclaration & param.CodeString & ", "
    Next param
    If mParameters.count > 0 Then strDeclaration = Left(strDeclaration, Len(strDeclaration) - 2)
    strDeclaration = strDeclaration & ")"
    If VarTypeString <> "" Then strDeclaration = strDeclaration & " As " & VarTypeString
    
    DeclarationString = strDeclaration
End Function

Public Function CodeString(Optional ByVal strPrefix As String = "") As String
    Dim strCode As String
    strCode = DeclarationString & vbCrLf
    strCode = strCode & mstrCode
    strCode = strCode & "End " & FunctionTypeString(blnTrailingSpace:=False, blnEnd:=True)
    
    CodeString = strCode
End Function




Public Property Get FunctionType() As enmFunctionType
    FunctionType = mFunctionType
End Property
Public Property Let FunctionType(ByVal val As enmFunctionType)
    mFunctionType = val
End Property

Private Function FunctionTypeString(Optional ByVal blnTrailingSpace As Boolean = True, _
                                    Optional ByVal blnEnd As Boolean = False) As String
    Dim strSpace As String
    If blnTrailingSpace Then strSpace = " "
    If blnEnd And Me.IsProperty Then
        FunctionTypeString = "Property" & strSpace
        Exit Function
    End If
    
    FunctionTypeString = CodeEnums.GetFunctionTypeString(Me.FunctionType) & strSpace
End Function

Public Property Get AccessLevel() As enmAccessLevel
    AccessLevel = mAccessLevel
End Property
Public Property Let AccessLevel(ByVal val As enmAccessLevel)
    mAccessLevel = val
End Property

Private Function AccessLevelString() As String
    'If Me.AccessLevel = acsBlank Then AccessLevelString = "": Exit Function
    If Me.AccessLevel <> acsBlank Then AccessLevelString = CodeEnums.GetAccessLevelString(Me.AccessLevel) & " "
End Function

Public Property Get VarType() As enmVarType
    VarType = mVarType
End Property
Public Property Let VarType(ByVal val As enmVarType)
    mVarType = val
End Property

Private Property Get VarTypeString() As String
    If Me.FunctionType = fncSub Or Me.FunctionType = fncPropertyLet Or Me.FunctionType = fncPropertySet Then
        VarTypeString = ""
    ElseIf Me.VarType = varOverride Then
        VarTypeString = mstrVarTypeOverride
    Else
        VarTypeString = CodeEnums.GetVarTypeString(Me.VarType)
    End If
End Property
Private Property Let VarTypeString(ByVal val As String)
    Dim vType As enmVarType
    vType = CodeEnums.GetVarTypeEnum(val)
    If vType = -1 Then
        Me.VarType = varOverride
        mstrVarTypeOverride = val
    Else
        Me.VarType = vType
    End If
End Property

'Public Property Get Parameters() As ParamType
'    Parameters = mParameters
'End Property
'Private Property Let Parameters(ByVal val As ParamType)
'    mParameters = val
'End Property

'Public Property Get Code() As String
'    Code = mstrCode
'End Property
'Private Property Let Code(ByVal val As String)
'    mstrCode = val
'End Property

Public Property Get Name() As String
    Name = mstrName
End Property
Public Property Let Name(ByVal val As String)
    mstrName = val
End Property

Public Property Get IsProperty() As Boolean
    Select Case Me.FunctionType
        Case fncPropertyGet
            IsProperty = True
        Case fncPropertyLet
            IsProperty = True
        Case fncPropertySet
            IsProperty = True
        Case Else
            IsProperty = False
    End Select
End Property
