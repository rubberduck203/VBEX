Attribute VB_Name = "TestList"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
'
' Constructors
' ------------
'
'@TestMethod
Public Sub ListCopy()

    Assert.AreEqual "List(1, 2, 3)", List.Copy(Array(1, 2, 3)).Show
    Assert.AreEqual "List()", List.Copy(Array()).Show
    
End Sub
'@TestMethod
Public Sub ListCopyIsCopy()

    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    
    Dim ys As List
    Set ys = List.Copy(xs)
    
    Assert.IsTrue xs.Equals(ys)
    Assert.AreNotEqual ObjPtr(xs), ObjPtr(ys)
    
End Sub
'@TestMethod
Public Sub ListNested()

    Dim flat As List
    Set flat = List.Create(1, 2, 3)
    
    Dim nested As List
    Set nested = List.Create(flat)
    
    Dim nestedCopy As List
    Set nestedCopy = List.Create(List.Copy(flat))
    
    flat.append 4
    
    Dim newNested As List
    Set newNested = List.Create(flat)
    
    Assert.IsTrue newNested.Equals(nested)
    Assert.IsFalse nested.Equals(nestedCopy)
    
End Sub
'@TestMethod
Public Sub ListRepeat()

    Dim xs As List
    Set xs = List.Repeat("x", 5)
    
    Assert.AreEqual CLng(5), xs.Count
    
    Dim x
    For Each x In xs
        Assert.AreEqual "x", x
    Next
    
End Sub
'
' Interfaces
' ----------
'
' ### IEquatable
'
'@TestMethod
Public Sub ListEquals()

    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    
    Assert.IsTrue xs.Equals(xs), "Self is equal to self"
    Assert.IsTrue xs.Equals(List.Create(1, 2, 3)), "equal to new instance"
    Assert.IsTrue List.Create(1, 2, 3).Equals(xs), "New Instance is equal to xs"
    
End Sub
'@TestMethod
Public Sub ListNotEquals()

    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    
    Assert.IsFalse xs.Equals(List.Create(4, 5, 6)), "same size, different elements"
    Assert.IsFalse xs.Equals(List.Create(1, 2, 3, 4)), "different size, same elements"
    Assert.IsFalse xs.Equals(List.Create("A", "B", "C")), "different element type"
    
End Sub
'
' ### ICountable
'
'@TestMethod
Public Sub ListCount()
    Assert.AreEqual List.Create(1, 2, 3).Count, CLng(3), "NonEmpty"
    Assert.AreEqual List.Create().Count, CLng(0), "empty"
End Sub
'
' ### ISequence
'
'@TestMethod
Public Sub ListToArray()
    Assert.AreEqual Join(Array(1, 2, 3)), Join(List.Create(1, 2, 3).ToArray), "multiple elements"
    Assert.IsNothing Join(List.Create().ToArray)
End Sub
'
' ### IPrintable
'
'@TestMethod
Public Sub ListShow()

    Dim flatList As List
    Set flatList = List.Create(1, 2, 3)
    
    Dim nestList As List
    Set nestList = List.Create(flatList, flatList)
    
    With Assert
        .AreEqual "List()", List.Create().Show
        .AreEqual "List(1, 2, 3)", flatList.Show
        .AreEqual "List(List(1, 2, 3), List(1, 2, 3))", nestList.Show
    End With
    
End Sub
'
' Methods
' -------
'
'@TestMethod
'Public Sub ListReduce()
'
'    Assert.AreEqual "abc", List.Create("a", "b", "c") _
'        .Reduce(Lambda.FromString("(a, b) => a & b"))
'
'End Sub
'@TestMethod
Public Sub ListFold()

    Assert.AreEqual "abc", List.Create("a", "b", "c") _
        .Fold("", Lambda.FromProper("(a, b) => a & b"))
        
End Sub

