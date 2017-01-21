VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements XMLObject


Private mColor As String
Private mLength As Integer


Public Property Get XmlObject_Color() As String
    XmlObject_Color = mColor
End Property
Public Property Let XmlObject_Color(ByVal val As String)
    mColor = val
End Property

Public Property Get XmlObject_Length() As Integer
    XmlObject_Length = mLength
End Property
Public Property Let XmlObject_Length(ByVal val As Integer)
    mLength = val
End Property


