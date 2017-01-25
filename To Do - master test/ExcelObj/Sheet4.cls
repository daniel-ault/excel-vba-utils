'Private Sub Worksheet_Activate()
'    Call RefreshSettings
'    'ComboBox1.List = GetCategoriesListArray()
'End Sub


Private Sub RefreshSettings()
    Worksheets("Settings").Range("tblSettingsStages[Stages]").RemoveDuplicates
End Sub
