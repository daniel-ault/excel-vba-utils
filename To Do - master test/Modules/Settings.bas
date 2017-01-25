Attribute VB_Name = "Settings"


Public mNumCheckBoxes As Integer
Public chkBoxes(1 To 10) As Boolean
Public chkBoxesId(1 To 10) As String

Public mdctChkBoxes As Scripting.Dictionary
Public mChkBoxes() As Checkbox
Public mCategories As Scripting.Dictionary
Public mintCategoryPage As Integer
Public mintCategoryNumPages As Integer

Public mToggleButton As toggleButton

Public mRibbon As IRibbonUI

Private mSettingsInitialized As Boolean

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public mChk As IRibbonControl


Public Sub Settings_Initialize()
    mNumCheckBoxes = 9
    mintCategoryPage = 1
    
    Set mdctChkBoxes = New Scripting.Dictionary
    Set mCategories = New Scripting.Dictionary
    ReDim mChkBoxes(mNumCheckBoxes) As Checkbox
    'Set mRibbon = New IRibbonUI
    
    CategoriesInit
    ChkBoxInit
    
    
    mintCategoryNumPages = Int(Ceiling(mCategories.count / UBound(mChkBoxes, 1)))
    
    Ribbon_Invalidate
    
    mSettingsInitialized = True
End Sub

Public Sub test()
    Dim test As Integer
    Dim testCount As Integer
    Dim testPages As Integer
    
    testCount = 23
    test = 10
    
    testPages = testCount - Int(23 / 10)
    
    MsgBox Ceiling(testCount / test)
End Sub

Sub onLoadRibbon_Custom(ByVal ribbon As Office.IRibbonUI)
    Set mRibbon = ribbon
    
    Worksheets("Settings Main").Range("H6").Value = VarPtr(ribbon)
End Sub

Sub Ribbon_Invalidate()
    mRibbon.Invalidate
End Sub

Private Sub ChkBoxInit()
    Dim i As Integer
    Set mdctChkBoxes = New Scripting.Dictionary
    
    For i = 1 To mNumCheckBoxes
        Set mChkBoxes(i) = New Checkbox
    Next i
    
    Set mToggleButton = New toggleButton
    
    Dim intStart, intEnd As Integer
    intEnd = mintCategoryPage * mNumCheckBoxes
    intStart = intEnd - mNumCheckBoxes + 1
    
    'If intEnd > mNumCheckBoxes Then intEnd = mNumCheckBoxes
    
    For i = 1 To mNumCheckBoxes
    'For i = intStart To intEnd
        'Dim chkBox As New Checkbox
        mChkBoxes(i).Id = "checkBox" & i
        mChkBoxes(i).Pressed = False
        'mChkBoxes(i).Label = mCategories.Keys((i + intStart - 2) - 1)
        mChkBoxes(i).Label = mCategories.Keys(i - 1)
        mdctChkBoxes.Add mChkBoxes(i).Id, i
        'Set mChkBoxes(i) = chkBox
    Next i
    
    mToggleButton.Id = "toggleButton1"
    mToggleButton.Pressed = False
    mToggleButton.Label = "Exclude mode"
    
'    For i = LBound(chkBoxesId) To UBound(chkBoxesId)
'        chkBoxesId(i) = "checkBox" & i
'    Next i
End Sub

Public Sub chkBoxes1()
    Dim i As Integer
    i = 0
    
    For i = 1 To UBound(chkBoxes)
        chkBoxes(i) = False
    Next i
    
    chkBoxes(1) = True
    chkBoxes(8) = True
End Sub

Public Sub CategoriesInit()
    'SetCategoriesList
    Dim strList As String
    strList = "Later"
    
    Dim arrCategories As Scripting.Dictionary, data As Variant
    Dim i As Integer
    'arrCategories = Worksheets("Settings (2)").Range("Table46[Categories]").value
    data = Worksheets(strList).Range("Table610[Category]").Value
    Set arrCategories = New Scripting.Dictionary
    
    For i = 1 To UBound(data)
        arrCategories(data(i, 1)) = Empty
    Next i
    
    Set mCategories = arrCategories
    
    'arrCategories =
    
'    Dim o As Variant
'    For Each o In arrCategories.Keys
'        MsgBox o
'    Next o
    
'    Dim i As Integer
'    For i = LBound(arrCategories) To UBound(arrCategories)
'        MsgBox arrCategories(i)
'    Next i
End Sub

Private Sub SetCategoriesList()
    Dim shSrc As Worksheet
    Dim shDest As Worksheet
    Dim rCategories As Range
    
    Dim strList As String
    strList = "Later"
    
    Set shSrc = Worksheets(strList)
    Set shDest = Worksheets("Settings (2)")
    
    Set rCategories = shDest.Range("Table46[Categories]")
    
    Worksheets(strList).Range ("Table610[Category]")
    rCategories.PasteSpecial
    
    rCategories.RemoveDuplicates (1)
    
    Application.CutCopyMode = False
    'Cells(1, 1).Select
End Sub

Public Sub chkBoxes2()
    Dim i As Integer
    i = 0
    
    For i = 1 To UBound(chkBoxes)
        chkBoxes(i) = False
    Next i
    
    chkBoxes(3) = True
    chkBoxes(4) = True
End Sub

Sub GetChkBoxLabel(control As IRibbonControl, ByRef Label)
'    Select Case control.Id
'        Case "checkBox1":
'            Label = "testing really really long name"
'            'mChk1 = control
'        Case "checkBox2":
'            Label = "wait whaaaaaat 72 no way"
'            'mChk2 = control
'        Case Else:
'            Label = control.Id
'    End Select
    
    If Not mSettingsInitialized Then Settings_Initialize
    
    If control.Id = "toggleButton1" Then
        Dim toggleBtn As toggleButton
        Set toggleBtn = mToggleButton
        Label = toggleBtn.Label
        Exit Sub
    End If

    Dim chkBox As Checkbox
    Set chkBox = mChkBoxes(mdctChkBoxes.Item(control.Id))
    Label = chkBox.Label
End Sub

Sub GetChkBoxPressed(control As IRibbonControl, ByRef bolReturn)
    Dim i As Integer
    
    If Not mSettingsInitialized Then Settings_Initialize
    
    'mChkBoxes (mdctChkBoxes(control.Id))

'    For i = LBound(chkBoxes) To UBound(chkBoxes)
'        If control.Id = chkBoxesId(i) Then
'            bolReturn = chkBoxes(i)
'
'        End If
'    Next i
End Sub

Sub checkBoxAction(control As IRibbonControl, Pressed As Boolean)
    If Not mSettingsInitialized Then Settings_Initialize
    
    If control.Id = "toggleButton1" Then
        Dim toggleBtn As toggleButton
        Set toggleBtn = mToggleButton
        mToggleButton.Pressed = Pressed
        
        SetFilters
        Exit Sub
    End If
    
    Dim chkBox As Checkbox
    Set chkBox = mChkBoxes(mdctChkBoxes(control.Id))
    mChkBoxes(mdctChkBoxes(control.Id)).Pressed = Pressed
    
    SetFilters
    'MsgBox chkBox.Label & " - " & Pressed
End Sub

Sub buttonAction(control As IRibbonControl)
    If control.Id = "buttonNext" And mintCategoryPage <> mintCategoryNumPages Then
        mintCategoryPage = mintCategoryPage + 1
    ElseIf control.Id = "buttonPrevious" And mintCategoryPage <> 1 Then
        mintCategoryPage = mintCategoryPage - 1
    End If
    
    ChkBoxInit
    Ribbon_Invalidate
End Sub

Sub SetFilters()
    Dim strTable As String
    strTable = "Table610"
    
    Dim blnIncludeMode As Boolean
    blnIncludeMode = Not mToggleButton.Pressed
    
    Dim i, nFilterCount As Integer
    nFilterCount = 0
    
    For i = 1 To UBound(mChkBoxes)
        If mChkBoxes(i).Pressed = blnIncludeMode Then nFilterCount = nFilterCount + 1
    Next i
    
    Dim tbl As Range
    Set tbl = Application.Range(strTable)
    
    Dim intCol As Integer
    intCol = Application.Range(strTable & "[Category]").Column
    intCol = intCol - tbl.Column + 1
    
    If nFilterCount = 0 Then
        tbl.AutoFilter Field:=intCol
        tbl.Worksheet.Range("I2").Calculate
        Exit Sub
    End If
    
    Dim arrFilters() As String
    ReDim arrFilters(nFilterCount) As String
    
    Dim j As Integer
    j = 1
    
    For i = 1 To UBound(mChkBoxes)
        If mChkBoxes(i).Pressed = blnIncludeMode Then
            arrFilters(j) = mChkBoxes(i).Label
            j = j + 1
        End If
    Next i
    
    'tbl.AutoFilter Field:=intCol
    tbl.AutoFilter Field:=intCol, Criteria1:=arrFilters, Operator:=xlFilterValues
    tbl.Worksheet.Range("I2").Calculate
    'tbl.Worksheet.AutoFilter.ApplyFilter
End Sub

Public Sub RefreshSettings()
    
End Sub

Public Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

'Public Sub test()
'    MsgBox UBound(chkBoxes) & " " & LBound(chkBoxes)
'End Sub
