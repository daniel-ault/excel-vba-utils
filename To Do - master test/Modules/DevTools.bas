Attribute VB_Name = "DevTools"
'Public Sub ExportSourceFiles()
'    Dim strPath As String
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    strPath = ActiveWorkbook.path & "\Exported Code\"
'    If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
'    strPath = strPath & RemoveExtension(ActiveWorkbook.Name) & "\"
'    If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
'
'    ExportSourceFilesToPath strPath
'End Sub
'
'Public Sub ImportSourceFiles()
'    Dim strPath As String
'    strPath = ActiveWorkbook.path & "\Exported Code\" & RemoveExtension(ActiveWorkbook.Name)
'
'    ImportSourceFilesFromPath strPath
'End Sub
'
'Public Sub ImportSourceFilesFromPath(sourcePath As String)
'    Dim file As String
'    file = Dir(sourcePath)
'
'    While (file <> vbNullString)
'        Application.VBE.ActiveVBProject.vbComponents.Import sourcePath & file
'        file = Dir
'    Wend
'End Sub
'
'Public Sub ExportSourceFilesToPath(destPath As String)
'    Dim component As VBIDE.vbComponent
'
'    For Each component In Application.VBE.ActiveVBProject.vbComponents
'        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
'            component.Export destPath & component.Name & ToFileExtension(component.Type)
'        End If
'    Next
'End Sub
'
''Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
''    Select Case vbeComponentType
''        Case vbext_ComponentType.vbext_ct_ClassModule
''            ToFileExtension = ".cls"
''        Case vbext_ComponentType.vbext_ct_StdModule
''            ToFileExtension = ".bas"
''        Case vbext_ComponentType.vbext_ct_MSForm
''            ToFileExtension = ".frm"
''        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
''        Case vbext_ComponentType.vbext_ct_Document
''        Case Else
''            ToFileExtension = vbNullString
''    End Select
''End Function
'
'Public Sub RemoveAllModules()
'    Dim project As VBProject
'    Set project = Application.VBE.ActiveVBProject
'
'    Dim comp As vbComponent
'    For Each comp In project.vbComponents
'        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
'            project.vbComponents.Remove comp
'        End If
'    Next
'End Sub
'
'Public Function RemoveExtension(ByVal strFileName As String)
'    Dim arr As Variant
'    arr = Split(ActiveWorkbook.Name, ".")
'    arr(UBound(arr)) = ""
'    RemoveExtension = Join(arr)
'End Function
