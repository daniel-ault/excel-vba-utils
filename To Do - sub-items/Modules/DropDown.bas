Attribute VB_Name = "DropDown"


Public Sub IsDropEnabled(control As IRibbonControl, ByRef Enabled)
    Enabled = True
End Sub


Public Sub getItemCount(control As IRibbonControl, ByRef count)
    count = Worksheets("Settings").Range("tblSettingsStages[Stages]").Rows.count
End Sub


Public Sub getItemLabel(control As IRibbonControl, index As Integer, ByRef Label)
    Label = Worksheets("Settings").Range("tblSettingsStages").Cells(index + 1).Value
    'label = index & " " & label
End Sub


Public Sub getLabel(control As IRibbonControl, ByRef Label)
    Label = "Stages"
End Sub


Public Sub onAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    MsgBox selectedId
End Sub

