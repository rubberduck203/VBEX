Attribute VB_Name = "IteratorTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

'@TestMethod
Public Sub CreateIterator()

    Dim xs As List
    Set xs = List.Create
    
    Dim it As Iterator
    Set it = Iterator.Create(xs)
    
    Assert.AreEqual "Iterator(0, List())", it.ToString

End Sub
'@TestMethod
Public Sub IncrementIterator()

    Dim xs As List
    Set xs = List.Create(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    
    Dim it As Iterator
    Set it = Iterator.Create(xs)
    
    Dim i As Integer
    i = 1
    Do While it.Inc
        Assert.AreEqual i, it.DeRef
        i = i + 1
    Loop
End Sub
