Attribute VB_Name = "TestShow"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub ShowPrimativeDataTest()

    With Assert
        .areequal "x", defshow.Show("x")
        .areequal "1", defshow.Show(CInt(1))
        .areequal "1", defshow.Show(CDbl(1))
    End With

End Sub
'@TestMethod
Public Sub ShowableTest()

    Dim flatList As List
    Set flatList = List.Create(1, 2, 3)
    
    Dim nestList As List
    Set nestList = List.Create(flatList, flatList)

    With Assert
        .areequal flatList.Show, defshow.Show(flatList)
        .areequal nestList.Show, defshow.Show(nestList)
    End With

End Sub
'@TestMethod
Public Sub ShowObjectTest()

    Dim c As New Collection
    Assert.IsTrue (defshow.Show(c) Like "Collection(&*)")

End Sub
