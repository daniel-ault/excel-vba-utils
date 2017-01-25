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
    
    
'    MsgBox "Do you want to commit changes?", vbDefaultButton1
'    Dim strCommitMessage As String
'    strCommitMessage = InputBox("Enter in commit message:")
    GitCommit ' strCommitMessage
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
'    Dim result As VbMsgBoxResult
'    Dim strCommitMessage As String
'
'    result = MsgBox("Do you want to commit changes?", vbYesNo)
'    If result = VbMsgBoxResult.vbNo Then Exit Sub
'
'    strCommitMessage = Application.InputBox("Enter commit message:")
'    If strCommitMessage = "False" Then Exit Sub
'    'commit code here
'
'    result = MsgBox("Do you want to push changes to GitHub?", vbYesNo)
'    If result = VbMsgBoxResult.vbNo Then Exit Sub
'
'    MsgBox strCommitMessage

    Dim strPath As String: strPath = "C:\Users\Daniel\Documents\Programming\Excel Utils\Exported Code"
    Dim strGit As String: strGit = "git -C """ & strPath & """"
    
    'MsgBox ShellRun(strGit & " status")
End Sub

Sub GitCommit(Optional ByVal strCommitMessage = "")
    Dim strPath As String: strPath = "C:\Users\Daniel\Documents\Programming\Excel Utils\Exported Code"
    Dim strGitPath As String: strGitPath = "C:\Users\Daniel\Documents\Programming\PortableGit\"
    
    'Dim strGit As String: strGit = """" & strGitPath & "git-bash.exe"" -C """ & strPath & """"
    Dim strGit As String: strGit = "git -C """ & strPath & """"
    
    Dim result As VbMsgBoxResult
    
    result = MsgBox("Do you want to commit changes?", vbYesNo)
    If result = VbMsgBoxResult.vbNo Then Exit Sub
    
    
    strCommitMessage = Application.InputBox("Enter commit message:")
    If strCommitMessage = "False" Then Exit Sub
    'commit code here
    MsgBox ShellRun(strGit & " add --all")
    Sleep 500
    MsgBox ShellRun(strGit & " commit -m """ & strCommitMessage & """")
    Sleep 500
    
    result = MsgBox("Do you want to push changes to GitHub?", vbYesNo)
    If result = VbMsgBoxResult.vbNo Then Exit Sub
    
'    If strCommitMessage = "" Then
'        strCommitMessage = "Automatically Committed via VBA whoa"
'    End If
    'Sleep 500
    Shell strGit & " push origin master", vbNormalFocus
End Sub


'taken from bburns.km on StackOverflow
'http://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba
Public Function ShellRun(sCmd As String) As String
    'Run a shell command, returning the output as a string'

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command'
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        'If sLine <> "" Then
        s = s & sLine & vbCrLf
    Wend

    ShellRun = s
End Function

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
    
    Dim oFolder As Object
    Dim oFile As Object
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
