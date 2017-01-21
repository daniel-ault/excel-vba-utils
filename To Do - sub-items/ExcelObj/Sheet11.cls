'Private Sub Worksheet_Activate()
'    Call HideSheets
'End Sub


Private Sub HideSheets()
    Dim sht As Worksheet
    Dim i As Integer
    i = 1
    
    For Each sht In Excel.ActiveWorkbook.Sheets
        If sht.Name Like "Settings*" And sht.Name <> Me.Name Then
            ActiveWorkbook.Sheets(sht.Name).Visible = xlSheetVisible 'xlSheetHidden
            Call CreateHyperlink(sht.Name, Me.Cells(i, 1).Address)
            i = i + 1
        End If
    Next sht
End Sub


Private Sub CreateHyperlink(strDest As String, strCell As String)
    Me.Hyperlinks.Add Anchor:=Me.Range(strCell), Address:="", SubAddress:="'" & strDest & "'!A1"
End Sub


Public Sub test()
    Sheet1.Select
End Sub



