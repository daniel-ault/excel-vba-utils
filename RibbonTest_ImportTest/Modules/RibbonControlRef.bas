Attribute VB_Name = "RibbonControlRef"

Public Sub RibbonControlRef_IsValidAttribute(attrType As AttributeType, ctlType As ControlType)
    MsgBox attrType & " - " & ctlType
End Sub

Public Sub RibbonControlRef_IsValidCallback(cbkType As CallbackType, ctlType As ControlType)
    MsgBox cbkType & " - " & ctlType
End Sub

Public Sub RibbonControlRef_IsValidChild(ctlChild As ControlType, ctlParent As ControlType)

End Sub

Public Sub RibbonControlRef_IsValidParent(ctlParent As ControlType, ctlChild As ControlType)

End Sub

