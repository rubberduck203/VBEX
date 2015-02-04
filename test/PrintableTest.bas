Attribute VB_Name = "PrintableTest"
'@TestModule
Option Explicit
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub PrintableToStringPrimativeData()

    With Assert
        .AreEqual "x", Printable.ToString("x")
        .AreEqual "1", Printable.ToString(CInt(1))
        .AreEqual "1", Printable.ToString(CDbl(1))
    End With

End Sub
'@TestMethod
Public Sub PrintableToStringIPrintable()

    Dim flatList As List
    Set flatList = List.Create(1, 2, 3)
    
    Dim nestList As List
    Set nestList = List.Create(flatList, flatList)

    With Assert
        .AreEqual flatList.ToString, Printable.ToString(flatList)
        .AreEqual nestList.ToString, Printable.ToString(nestList)
    End With

End Sub
'@TestMethod
Public Sub PrintableToStringObject()

    Dim c As New Collection
    Assert.IsTrue (Printable.ToString(c) Like "Collection(&*)")

End Sub
