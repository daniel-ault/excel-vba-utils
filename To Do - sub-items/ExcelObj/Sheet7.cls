'Private Sub Worksheet_Activate()
'    Call SetCategoriesList
'    'ComboBox1.List = GetCategoriesListArray()
'End Sub

Private Sub Worksheet_Deactivate()
    
End Sub



Private Sub SetCategoriesList()
    Dim shSrc As Worksheet
    Dim shDest As Worksheet
    Dim rCategories As Range
    
    Set shSrc = Worksheets("To Do List")
    Set shDest = Worksheets("Settings (2)")
    
    Set rCategories = shDest.Range("Table46[Categories]")
    
    Worksheets("To Do List").Range("Table6[Category]").Copy
    rCategories.PasteSpecial
    
    rCategories.RemoveDuplicates (1)
    
    Application.CutCopyMode = False
    Cells(1, 1).Select
End Sub

Public Function GetCategoriesListArray() As Variant
    Dim shSrc As Worksheet
    Dim shDest As Worksheet
    Dim rCategories As Range
    
    Set shSrc = Worksheets("To Do List")
    Set shDest = Worksheets("Settings (2)")
    
    Set rCategories = shDest.Range("Table46[Categories]")
    
    Dim vArr() As Variant
    'Dim arr() As String
    vArr = Worksheets("To Do List").Range("Table6[Category]").Value
    
    'vArr = DictionaryUnique(vArr)
    GetCategoriesListArray = DictionaryUnique(vArr)
    
    'shDest.Range("Table46[Categories]") = vArr
    'rCategories.Clear
    
    
'    For Each Item In vArr
'        MsgBox Item
'    Next Item
'
    'arr = DictionaryTest(vArr, 1000)
    
'    Dim i As Integer
'    i = 1
'
'    For Each Item In vArr
'        MsgBox Item
'        i = i + 1
'        If i = 10 Then Exit For
'    Next Item

    
'    For i = LBound(vArr) To 10 'UBound(arr)
'        MsgBox TypeName(vArr(2))
'    Next i
    
    'rCategories.RemoveDuplicates (1)
    
    'Application.CutCopyMode = False
    'Cells(1, 1).Select
End Function


Function DictionaryUnique(ByRef arr() As Variant) As Variant
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    For Each Item In arr
        If Not d.Exists(Item) Then
            d.Add Item, 1
        End If
    Next Item
    
    DictionaryUnique = d.Keys()
End Function

Function DictionaryTest(ByRef myArray() As Variant, Lim As Long) As Variant
    Dim StrtTime As Double, Endtime As Double
    Dim d As Scripting.Dictionary, i As Long  '' Early Binding
    Set d = New Scripting.Dictionary
    For i = LBound(myArray) To UBound(myArray): d(myArray(i)) = 1: Next i
    DictionaryTest = d.Keys()
End Function



