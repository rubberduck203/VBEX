Attribute VB_Name = "TestShow"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub ShowPrimativeDataTest()

    With Assert
        .AreEqual "x", defshow.Show("x")
        .AreEqual "1", defshow.Show(CInt(1))
        .AreEqual "1", defshow.Show(CDbl(1))
    End With

End Sub
'@TestMethod
Public Sub ShowableTest()

    Dim flatList As List
    Set flatList = List.Create(1, 2, 3)
    
    Dim nestList As List
    Set nestList = List.Create(flatList, flatList)

    With Assert
        .AreEqual flatList.Show, defshow.Show(flatList)
        .AreEqual nestList.Show, defshow.Show(nestList)
    End With

End Sub
'@TestMethod
Public Sub ShowObjectTest()

    Dim c As New Collection
    Assert.IsTrue (defshow.Show(c) Like "Collection(&*)")

End Sub
