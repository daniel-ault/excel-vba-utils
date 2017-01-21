Attribute VB_Name = "Functions"
Public Function GetTableName(ByVal rng As Range)
    GetTableName = rng.ListObject.Name
End Function


Public Sub AddHighlightRibbon()
    Dim ribbonXML As String
    
    ribbonXML = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    ribbonXML = ribbonXML + "  <mso:ribbon>"
    ribbonXML = ribbonXML + "    <mso:qat/>"
    ribbonXML = ribbonXML + "    <mso:tabs>"
    ribbonXML = ribbonXML + "      <mso:tab id=""highlightTab"" label=""Highlight"" insertBeforeQ=""mso:TabFormat"">"
    ribbonXML = ribbonXML + "        <mso:group id=""testGroup"" label=""Test"" autoScale=""true"">"
    ribbonXML = ribbonXML + "          <mso:button id=""highlightManualTasks"" label=""Toggle Manual Task Color"" "
    ribbonXML = ribbonXML + "imageMso=""DiagramTargetInsertClassic"" onAction=""ToggleManualTasksColor""/>"
    ribbonXML = ribbonXML + "        </mso:group>"
    ribbonXML = ribbonXML + "      </mso:tab>"
    ribbonXML = ribbonXML + "    </mso:tabs>"
    ribbonXML = ribbonXML + "  </mso:ribbon>"
    ribbonXML = ribbonXML + "</mso:customUI>"
    
    ActiveProject.SetCustomUI (ribbonXML)
End Sub



Public Sub StringToColumn(str As String)
    
End Sub
