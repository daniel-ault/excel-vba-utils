Attribute VB_Name = "Testing"

Public Sub test()
    'Dim testobj As New XMLObject
    
    Dim testtest As XMLObject
    Set testtest = New XmlTest
'
'    Dim testtest2 As XMLObject
'    Set testtest2 = New XmlTest2

    Dim testtest3 As XMLObject
    Set testtest3 = New XmlTest3
    
    testtest.Color = "blue"
'    testtest2.Color = "five"
    testtest3.Color = "banana"
    'testtest3.Shape = "whoa"
    
    'MsgBox testtest3.Shape
    
    'MsgBox testtest.Color
    'MsgBox testtest2.Color
    
'    MsgBox TypeName(testtest)
    
End Sub


Public Sub test2()


    Dim tbl As Range
    Set tbl = Application.Range("Table610")
    
    Dim lngTimer As Long
    
    Dim sw As StopWatch
    Set sw = New StopWatch
    
    sw.StartTimer
    Dim arr As Variant
    arr = tbl.Value
    lngTimer = sw.EndTimer
    
    MsgBox lngTimer
    
    'MsgBox tbl.Address
End Sub

Public Sub test3()
    Dim strTest As String
    strTest = "123456"
End Sub
