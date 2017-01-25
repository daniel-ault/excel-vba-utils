VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mAccessLevel As enmAccessLevel
Private mVarType As enmVarType
Private mstrName As String
Private mblnIsParameter As Boolean
Private mParamType As enmParamType
Private mblnIsOptional As Boolean

Private mblnVarTypeOverride As Boolean
Private mstrVarTypeOverride As String

'TODO variable for member var, parameter, or sub

Public Sub init(ByVal strName As String, _
                ByVal vType As enmVarType, _
                Optional ByVal level As enmAccessLevel = acsBlank, _
                Optional ByVal blnIsParam As Boolean = False, _
                Optional ByVal blnIsOptional As Boolean = False, _
                Optional ByVal pParamType As enmParamType = prmByVal, _
                Optional ByVal strVarTypeOverride As String)
    Me.Name = strName
    If strVarTypeOverride <> "" Then
        VarTypeString = strVarTypeOverride
    Else
        Me.VarType = vType
    End If
    Me.AccessLevel = level
    Me.IsParameter = blnIsParam
    Me.ParamType = pParamType
    Me.IsOptional = blnIsOptional
End Sub


Public Function CodeString(Optional ByVal strPrefix As String = "") As String
    Dim str As String
    str = AccessLevelString
    str = str & IsOptionalString
    str = str & ParamTypeString
    str = str & strPrefix & Me.Name
    str = str & " As " & Me.VarTypeString
    
    CodeString = str
End Function

Public Property Get AccessLevel() As enmAccessLevel
    If Me.IsParameter Then
        AccessLevel = acsBlank
    Else
        AccessLevel = mAccessLevel
    End If
End Property
Public Property Let AccessLevel(ByVal val As enmAccessLevel)
    mAccessLevel = val
End Property

Private Function AccessLevelString() As String
    If Me.AccessLevel <> acsBlank Then AccessLevelString = CodeEnums.GetAccessLevelString(Me.AccessLevel) & " "
End Function

Public Property Get VarType() As enmVarType
    VarType = mVarType
End Property
Public Property Let VarType(ByVal val As enmVarType)
    mVarType = val
End Property

Public Property Get VarTypeString() As String
    If Me.VarType = varOverride Then
        VarTypeString = mstrVarTypeOverride
    Else
        VarTypeString = CodeEnums.GetVarTypeString(Me.VarType)
    End If
End Property
Public Property Let VarTypeString(ByVal val As String)
    Dim vType As enmVarType
    vType = CodeEnums.GetVarTypeEnum(val)
    If vType = -1 Then
        Me.VarType = varOverride
        mstrVarTypeOverride = val
    Else
        Me.VarType = vType
    End If
End Property

Public Property Get Name() As String
    Name = mstrName
End Property
Public Property Let Name(ByVal val As String)
    mstrName = val
End Property

Public Property Get IsParameter() As Boolean
    IsParameter = mblnIsParameter
End Property
Public Property Let IsParameter(ByVal val As Boolean)
    mblnIsParameter = val
End Property

Public Property Get IsOptional() As Boolean
    If Not Me.IsParameter Then
        IsOptional = False
    Else
        IsOptional = mblnIsOptional
    End If
End Property
Public Property Let IsOptional(ByVal val As Boolean)
    mblnIsOptional = val
End Property

Private Function IsOptionalString() As String
    If Me.IsOptional Then IsOptionalString = CodeEnums.GetOptionalString(True) & " "
End Function

Public Property Get ParamType() As enmParamType
    If Not Me.IsParameter Then
        ParamType = prmBlank
    Else
        ParamType = mParamType
    End If
End Property
Public Property Let ParamType(ByVal val As enmParamType)
    mParamType = val
End Property

Private Function ParamTypeString() As String
    If Me.ParamType <> prmBlank Then ParamTypeString = CodeEnums.GetParamTypeString(Me.ParamType) & " "
End Function
