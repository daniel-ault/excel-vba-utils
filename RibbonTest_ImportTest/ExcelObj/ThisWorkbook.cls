Private Sub Workbook_Activate()

'Dim hFile As Long
'Dim path As String, fileName As String, ribbonXML As String, user As String
'
'hFile = FreeFile
'user = Environ("Username")
'path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
'fileName = "Excel.officeUI"
'
'ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
'ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
'ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
'ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
'ribbonXML = ribbonXML + "      <mso:tab id='reportTab' label='My Actions' insertBeforeQ='mso:TabHome'>" & vbNewLine
'ribbonXML = ribbonXML + "        <mso:group id='reportGroup' label='Reports' autoScale='true'>" & vbNewLine
'ribbonXML = ribbonXML + "          <mso:button id='runReport' label='Trim' " & vbNewLine
'ribbonXML = ribbonXML + "imageMso='AppointmentColor3'      onAction='TrimSelection'/>" & vbNewLine
'ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
'ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
'ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
'ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
'ribbonXML = ribbonXML + "</mso:customUI>"
'
'ribbonXML = Replace(ribbonXML, """", "")
'
'Open path & fileName For Output Access Write As hFile
'Print #hFile, ribbonXML
'Close hFile

End Sub

Private Sub Workbook_Deactivate()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "Excel.officeUI"

ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
"<mso:ribbon></mso:ribbon></mso:customUI>"

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile

End Sub


Sub AddSht_AddCode()
    Dim wb As Workbook
    Dim xPro As VBIDE.VBProject
    Dim xCom As VBIDE.vbComponent
    Dim xMod As VBIDE.CodeModule
    Dim xLine As Long

    'Set wb = Workbooks.Add
    Set wb = Me

    With wb
        Set xPro = .VBProject
        Set xCom = xPro.vbComponents("Sheet1")
        Set xMod = xCom.CodeModule
        
        With xMod
            xLine = .CreateEventProc("Change", "Worksheet")
            xLine = xLine + 1
            .InsertLines xLine, "  Cells.Columns.AutoFit"
        End With
    End With

End Sub
