Attribute VB_Name = "Enums"
Enum ControlType
    ctlButton
    ctlCheckBox
    ctlToggleButton
    ctlDialogBoxLauncher
    ctlItem
    ctlBox
    ctlButtonGroup
    ctlGroup
    ctlLabelControl
    ctlSeparator
    ctlTab
    ctlComboBox
    ctlDropDown
    ctlDynamicMenu
    ctlEditBox
    ctlGallery
    ctlMenu
    ctlSplitButton
    ctlMenuSeparator
    ctlOfficeMenu
End Enum

Enum AttributeType
    attColumns
    attDescription
    attEnabled
    attId
    attIdMso
    attIdQ
    attImage
    attImageMso
    attInsertAfterMso
    attInsertAfterQ
    attInsertBeforeMso
    attInsertBeforeQ
    attInvalidateContentOnDrop
    attItemHeight
    attItemSize
    attItemWidth
    attKeytip
    attLabel
    attMaxLength
    attRows
    attScreentip
    attShowImage
    attShowItemImage
    attShowItemLabel
    attShowLabel
    attSize
    attSizeString
    attSupertip
    attTag
    attVisible
End Enum

Enum CallbackType
    clbGetContent
    clbGetDescription
    clbGetEnabled
    clbGetImage
    clbGetItemCount
    clbGetItemHeight
    clbGetItemID
    clbGetItemImage
    clbGetItemLabel
    clbGetItemScreentip
    clbGetItemSupertip
    clbGetItemWidth
    clbGetKeytip
    clbGetLabel
    clbGetPressed
    clbGetPressedChk
    clbGetScreentip
    clbGetSelectedItemID
    clbGetSelectedItemIndex
    clbGetShowImage
    clbGetShowLabel
    clbGetSize
    clbGetSupertip
    clbGetText
    clbGetTitle
    clbGetVisible
    clbOnAction
    clbOnActionChk
    clbOnChange
    clbOnActionBtn
    clbOnActionLst
End Enum


Private arrControlTypeStrings() As Variant
Private arrAttributeTypeStrings() As Variant
Private arrCallbackTypeStrings() As Variant

'Private arrControlToAttribute(numControls, numAttributes)

Private blnStringArraysSet As Boolean

Private Sub SetStringArrays()
    arrControlTypeStrings = Array("Button", "CheckBox", "ToggleButton", "DialogBoxLauncher", "Item", "Box", "ButtonGroup", "Group", "LabelControl", "Separator", "Tab", "ComboBox", "DropDown", "DynamicMenu", "EditBox", "Gallery", "Menu", "SplitButton", "MenuSeparator", "OfficeMenu")
    arrAttributeTypeStrings = Array("columns", "description", "enabled", "id", "idMso", "idQ", "image", "imageMso", "insertAfterMso", "insertAfterQ", "insertBeforeMso", "insertBeforeQ", "invalidateContentOnDrop", "itemHeight", "itemSize", "itemWidth", "keytip", "label", "maxLength", "rows", "screentip", "showImage", "showItemImage", "showItemLabel", "showLabel", "size", "sizeString", "supertip", "tag", "visible", "")
    arrCallbackTypeStrings = Array("getContent", "getDescription", "getEnabled", "getImage", "getItemCount", "getItemHeight", "getItemID", "getItemImage", "getItemLabel", "getItemScreentip", "getItemSupertip", "getItemWidth", "getKeytip", "getLabel", "getPressed", "getPressedChk", "getScreentip", "getSelectedItemID", "getSelectedItemIndex", "getShowImage", "getShowLabel", "getSize", "getSupertip", "getText", "getTitle", "getVisible", "onAction", "onActionChk", "onChange", "onActionBtn", "onActionLst", "")
    blnArraysSet = True
End Sub


Private Sub SetAssociationArrays()

End Sub

'Public Function IsValidAttribute(attType As AttributeType, ctlType As ControlType) As Boolean
'    If Not blnAssocArraysSet Then SetAssociationArrays
'
'    IsValidAttribute = arrControlToAttribute(ctlType, attType)
'End Function

Public Function GetControlTypeString(ctlControlType As ControlType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetControlTypeString = arrControlTypeStrings(ctlControlType)
End Function

Public Function GetAttributeTypeString(attAttributeType As AttributeType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetAttributeTypeString = arrAttributeTypeStrings(attAttributeType)
End Function

Public Function GetCallbackTypeString(clbCallbackType As CallbackType) As String
    If Not blnStringArraysSet Then SetStringArrays
    
    GetCallbackTypeString = arrCallbackTypeStrings(clbCallbackType)
End Function


