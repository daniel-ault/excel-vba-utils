Attribute VB_Name = "Checkbox"
Private mstrDescription As String
Private mblnEnabled As Boolean
Private mstrId As String
Private mstrIdMso As String
Private mstrIdQ As String
Private mstrInsertAfterMso As String
Private mstrInsertAfterQ As String
Private mstrInsertBeforeMso As String
Private mstrInsertBeforeQ As String
Private mstrKeytip As String
Private mstrLabel As String
Private mstrScreentip As String
Private mstrSupertip As String
Private mstrTag As String
Private mblnVisible As Boolean


Public Property Get Description() As String
    Description = mstrDescription
End Property
Public Property Let Description(ByVal val As String)
    mstrDescription = val
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
Public Property Let Enabled(ByVal val As Boolean)
    mblnEnabled = val
End Property

Public Property Get Id() As String
    Id = mstrId
End Property
Public Property Let Id(ByVal val As String)
    mstrId = val
End Property

Public Property Get IdMso() As String
    IdMso = mstrIdMso
End Property
Public Property Let IdMso(ByVal val As String)
    mstrIdMso = val
End Property

Public Property Get IdQ() As String
    IdQ = mstrIdQ
End Property
Public Property Let IdQ(ByVal val As String)
    mstrIdQ = val
End Property

Public Property Get InsertAfterMso() As String
    InsertAfterMso = mstrInsertAfterMso
End Property
Public Property Let InsertAfterMso(ByVal val As String)
    mstrInsertAfterMso = val
End Property

Public Property Get InsertAfterQ() As String
    InsertAfterQ = mstrInsertAfterQ
End Property
Public Property Let InsertAfterQ(ByVal val As String)
    mstrInsertAfterQ = val
End Property

Public Property Get InsertBeforeMso() As String
    InsertBeforeMso = mstrInsertBeforeMso
End Property
Public Property Let InsertBeforeMso(ByVal val As String)
    mstrInsertBeforeMso = val
End Property

Public Property Get InsertBeforeQ() As String
    InsertBeforeQ = mstrInsertBeforeQ
End Property
Public Property Let InsertBeforeQ(ByVal val As String)
    mstrInsertBeforeQ = val
End Property

Public Property Get Keytip() As String
    Keytip = mstrKeytip
End Property
Public Property Let Keytip(ByVal val As String)
    mstrKeytip = val
End Property

Public Property Get Label() As String
    Label = mstrLabel
End Property
Public Property Let Label(ByVal val As String)
    mstrLabel = val
End Property

Public Property Get Screentip() As String
    Screentip = mstrScreentip
End Property
Public Property Let Screentip(ByVal val As String)
    mstrScreentip = val
End Property

Public Property Get Supertip() As String
    Supertip = mstrSupertip
End Property
Public Property Let Supertip(ByVal val As String)
    mstrSupertip = val
End Property

Public Property Get Tag() As String
    Tag = mstrTag
End Property
Public Property Let Tag(ByVal val As String)
    mstrTag = val
End Property

Public Property Get Visible() As Boolean
    Visible = mblnVisible
End Property
Public Property Let Visible(ByVal val As Boolean)
    mblnVisible = val
End Property


