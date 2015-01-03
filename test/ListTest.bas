Attribute VB_Name = "ListTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass

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

'@TestMethod
Public Sub ListToArray()
    Assert.AreEqual Join(Array(1, 2, 3)), Join(List.Create(1, 2, 3).ToArray), "multiple elements"
    Assert.IsNothing Join(List.Create().ToArray)
End Sub

'@TestMethod
Public Sub ListToString()
    Assert.AreEqual "[1, 2, 3]", List.Create(1, 2, 3).ToString
    Assert.AreEqual "[]", List.Create().ToString
End Sub

'@TestMethod
Public Sub ListCount()
    Assert.AreEqual List.Create(1, 2, 3).Count, CLng(3), "NonEmpty"
    Assert.AreEqual List.Create().Count, CLng(0), "empty"
End Sub

'@TestMethod
Public Sub ListCopy()
    Assert.AreEqual "[1, 2, 3]", List.Copy(Array(1, 2, 3)).ToString
    Assert.AreEqual "[]", List.Copy(Array()).ToString
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
Public Sub ArrayToList()
    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    Dim ys As List
    Set ys = List.Copy(Array(1, 2, 3))
    Assert.IsTrue xs.Equals(ys)
    Assert.IsTrue ys.Equals(xs)
End Sub

'@TestMethod
Public Sub ListReduce()
    Assert.AreEqual "abc", List.Create("a", "b", "c") _
        .Reduce(Lambda.FromString("(a, b) => a & b"))
End Sub
'@TestMethod
Public Sub ListFold()
    Assert.AreEqual "abc", List.Create("a", "b", "c") _
        .Fold("", Lambda.FromString("(a, b) => a & b"))
End Sub

