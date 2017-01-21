Attribute VB_Name = "MacroButtons"

Option Explicit

Public mRibbon As IRibbonUI

'Private mChk1 As New IRibbonControl
'Private mChk2 As New IRibbonControl
Public mLabel As String

Public mTest As String

'Sub onLoadRibbon(ByVal ribbon As Office.IRibbonUI)
'
'    Set Settings.mRibbon = ribbon
'    Set mRibbon = ribbon
'
'    Worksheets("Settings Main").Range("H6").Value = VarPtr(ribbon)
'End Sub



Sub RefreshRibbon()
    'Set mRibbon = Settings.mRibbon
    'mRibbon.Invalidate
End Sub

'Callback for customButton1 onAction
Sub MacroButton1(control As IRibbonControl)
    'mRibbon.Invalidate
    'RibbonInvalidate
    'mTest = "It worked! :O"
End Sub

'Callback for customButton2 onAction
Sub MacroButton2(control As IRibbonControl)
    MsgBox mTest
    'MsgBox "This is MacroButton 2"
    'mRibbon.InvalidateControl
End Sub

'Callback for customButton3 onAction
Sub MacroButton3(control As IRibbonControl)
    MsgBox "This is MacroButton 3"
End Sub

'Callback for customButton4 onAction
Sub MacroButton4(control As IRibbonControl)
    'MsgBox "This is MacroButton 4"
    Call RefreshFilters
End Sub

'Callback for customButton5 onAction
Sub MacroButton5(control As IRibbonControl)
    'MsgBox "This is MacroButton 5"
    Call FinishRow
End Sub

'Callback for customButton6 onAction
Sub MacroButton6(control As IRibbonControl)
    'MsgBox "This is MacroButton 6"
    'UserForm1.Show
End Sub

'Callback for customButton7 onAction
Sub MacroButton7(control As IRibbonControl)
    'Call MoveToILOList
    Call MoveToNextStage
    Call RefreshFilters
End Sub

'Callback for customButton8 onAction
Sub MacroButton8(control As IRibbonControl)
    MsgBox "This is MacroButton 8"
End Sub

'Callback for customButton9 onAction
Sub MacroButton9(control As IRibbonControl)
    MsgBox "This is MacroButton 9"
End Sub

'Callback for customButton10 onAction
Sub MacroButton10(control As IRibbonControl)
    MsgBox "This is MacroButton 10"
End Sub

'Callback for customButton11 onAction
Sub MacroButton11(control As IRibbonControl)
    MsgBox "This is MacroButton 11"
End Sub

'Callback for customButton12 onAction
Sub MacroButton12(control As IRibbonControl)
    MsgBox "This is MacroButton 12"
End Sub





Sub TestGetSelItem()
    
End Sub

Sub TestOnAction()
    
End Sub

Public Sub TestGetItemCount()
    'TestGetItemCount = 5
End Sub

Sub TestGetItemID()
    
End Sub

Sub TestGetItemLabel()
    
End Sub
