Attribute VB_Name = "DevTools"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ExportSourceFiles()
    Dim strPath, strClassFolder, strModuleFolder, strExcelObjFolder As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strClassFolder = "Classes\"
    strModuleFolder = "Modules\"
    strExcelObjFolder = "ExcelObj\"
    
    Dim wb As Workbook
    Dim vbProj As VBProject
    Dim component As vbComponent
    
    For Each vbProj In Application.VBE.VBProjects
        strPath = MoveUpOneDir(vbProj.fileName) & "Exported Code\"
        CreateFolderNew strPath
        
        strPath = strPath & RemoveExtension(RemovePath(vbProj.fileName)) & "\"
        CreateFolderNew strPath
        
        CreateFolderNew strPath & strClassFolder
        CreateFolderNew strPath & strModuleFolder
        CreateFolderNew strPath & strExcelObjFolder
        
        'ExportSourceFilesToPath strPath, wb.Name
        For Each component In vbProj.vbComponents
            Select Case component.Type
                Case vbext_ComponentType.vbext_ct_ClassModule:
                    component.Export strPath & _
                                     strClassFolder & _
                                     component.Name & _
                                     ToFileExtension(component.Type)

                Case vbext_ComponentType.vbext_ct_StdModule:
                    component.Export strPath & _
                                     strModuleFolder & _
                                     component.Name & _
                                     ToFileExtension(component.Type)
                
                Case vbext_ComponentType.vbext_ct_Document:
                    ExportDocument component, _
                                   vbProj, _
                                   strPath & _
                                   strExcelObjFolder & _
                                   component.Name & _
                                   ToFileExtension(component.Type)
            End Select
'            If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
'                component.Export strPath & component.Name & ToFileExtension(component.Type)
'            End If
        Next
    Next vbProj
End Sub

Private Sub ExportDocument(vbComp As vbComponent, _
                           vbProj As VBProject, _
                           ByVal strFilePath As String)
    Dim src, dest As CodeModule
    Dim fso, oFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(strFilePath)
    Set src = vbComp.CodeModule
    
    If src.CountOfLines <> 0 Then
        oFile.WriteLine src.Lines(1, src.CountOfLines)
    End If
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing
    'Set dest = WriteCode.AddModule("ExcelObj_" & vbComp.Name, vbext_ct_StdModule)
    'WriteCode.AddCodeToModule "ExcelObj_" & vbComp.Name, src.Lines(1, src.CountOfLines)
End Sub


Public Sub testt()
    Dim vbComp As vbComponent
    
'    For Each vbComp In Application.VBE.ActiveVBProject.vbComponents
'        MsgBox vbComp.Name & " - " & vbComp.Type
'    Next vbComp

    MsgBox
End Sub

Sub GitCommit(Optional ByVal strCommitMessage = "")
    Dim strPath As String: strPath = "C:\Users\Daniel\Documents\Programming\Excel Utils\Exported Code"
    Dim strGit As String: strGit = "git -C """ & strPath & """"
    
    If strCommitMessage = "" Then
        strCommitMessage = "Automatically Committed via VBA whoa"
    End If
    
    Shell strGit & " add --all", vbNormalFocus
'    Shell strGit & " commit -m """ & strCommitMessage & """"
'    Shell strGit & " push origin master", vbNormalFocus
End Sub

Private Sub CreateFolderNew(ByVal strPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(strPath) Then
        fso.CreateFolder strPath
    Else
        DeleteFilesInFolder strPath
    End If
End Sub

Private Sub DeleteFilesInFolder(ByVal strPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim oFolder As Variant
    Dim oFile As Variant
    Set oFolder = fso.GetFolder(strPath)
    
    For Each oFile In oFolder.Files
        oFile.Delete True
    Next oFile
End Sub

Private Sub DeleteFilesAndFolder(ByVal strPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim oFolder As Folder
    Dim oFile As file
    Set oFolder = fso.GetFolder(strPath)
    
    For Each oFile In oFolder.Files
        oFile.Delete True
    Next oFile
    
    fso.DeleteFolder (strPath)
    
    'DeleteFilesAndFolder = oFolder.Files.count = 0
End Sub


Public Sub ImportSourceFiles()
    Dim strPath As String
    strPath = ActiveWorkbook.path & "\Exported Code\" & RemoveExtension(ActiveWorkbook.Name)
    
    ImportSourceFilesFromPath strPath
End Sub

Public Sub ImportSourceFilesFromPath(sourcePath As String, _
                                     Optional wb As Workbook)
    
    If IsNull(wb) Then Set wb = ActiveWorkbook
    
    Dim file As String
    file = Dir(sourcePath)
    
    While (file <> vbNullString)
        Application.VBE.ActiveVBProject.vbComponents.Import sourcePath & file
        file = Dir
    Wend
End Sub


 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
            ToFileExtension = ".cls"
        Case Else
            ToFileExtension = vbNullString
    End Select
End Function

Public Sub RemoveAllModules()
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
     
    Dim comp As vbComponent
    For Each comp In project.vbComponents
        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.vbComponents.Remove comp
        End If
    Next
End Sub

Public Function RemoveExtension(ByVal strFileName As String) As String
    Dim arr As Variant
    arr = Split(strFileName, ".")
    arr(UBound(arr)) = ""
    RemoveExtension = Join(arr, "")
End Function

Public Function RemovePath(ByVal strPath As String) As String
    Dim arr As Variant
    arr = Split(strPath, "\")
    RemovePath = arr(UBound(arr))
End Function

Public Function MoveUpOneDir(ByVal strPath As String) As String
    Dim arr As Variant
    arr = Split(strPath, "\")
    arr(UBound(arr)) = ""
    MoveUpOneDir = Join(arr, "\")
End Function
