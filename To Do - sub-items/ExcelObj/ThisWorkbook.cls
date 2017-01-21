Private Sub Workbook_Open()
    Settings_Initialize
End Sub


'Private Sub Workbook_SheetFollowHyperLink(ByVal Sh As Object, ByVal Target As Hyperlink)
'    Dim rng As Range
'    Set rng = Application.Evaluate(Target.SubAddress)
'
'    ActiveWorkbook.Sheets(rng.Parent.Name).Visible = xlSheetVisible
'    ActiveWorkbook.Sheets(rng.Parent.Name).Select
'End Sub





