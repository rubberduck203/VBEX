Attribute VB_Name = "SortedSetTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

Private Sub SetEmptyConstructorTest(ByVal emptyset As SortedSet)
    
    Assert.IsNotNothing emptyset, "Empty set is not nothing"
    Assert.AreEqual "SortedSet()", emptyset.ToString, "Emptyset repr SortedSet()"
    Assert.AreEqual CLng(0), emptyset.Count, "emptyset count = 0"
    
End Sub
Private Sub SetNonEmptyConstructorTest(ByVal nonempty As SortedSet)

    Assert.IsNotNothing nonempty, "nonempty set is not nothing"
    Assert.AreEqual "SortedSet(1, 2, 3)", nonempty.ToString, "nonempty repr SortedSet(1, 2, 3)"
    Assert.AreEqual CLng(3), nonempty.Count, "nonempty count = 0"

End Sub

'@TestMethod
Public Sub SetEmptyCreate()
    SetEmptyConstructorTest SortedSet.Create()
End Sub

'@TestMethod
Public Sub SetEmptyCopy()
    SetEmptyConstructorTest SortedSet.Copy(Array())
End Sub

'@TestMethod
Public Sub SetNonEmptyCopy()

    SetNonEmptyConstructorTest SortedSet.Copy(Array(1, 2, 3))
    SetNonEmptyConstructorTest SortedSet.Copy(Array(1, 2, 3, 1, 2, 3))
    
End Sub

'@TestMethod
Public Sub SetNonEmptyCreate()

    SetNonEmptyConstructorTest SortedSet.Create(1, 2, 3)
    SetNonEmptyConstructorTest SortedSet.Create(1, 2, 3, 1, 2, 3)
    
End Sub

'@TestMethod
Public Sub SetDifference()

    Dim S1 As SortedSet
    Set S1 = SortedSet.Create(1, 3, 5)
    
    Dim S2 As SortedSet
    Set S2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5, 7)
    
    Assert.AreEqual "SortedSet()", S1.Difference(S1).ToString
    Assert.AreEqual "SortedSet(1, 3, 5)", S1.Difference(S2).ToString
    Assert.AreEqual "SortedSet(2, 4, 6)", S2.Difference(S1).ToString
    
    Assert.AreEqual "SortedSet(7)", S3.Difference(S1).ToString
    Assert.AreEqual "SortedSet(1)", S1.Difference(S3).ToString
    
End Sub

Private Function Setup() As Tuple
    
    Dim S1 As SortedSet
    Set S1 = SortedSet.Create(1, 3, 5)
    
    Dim S2 As SortedSet
    Set S2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5, 7)
    
    Set Setup = Tuple.Pack(S1, S2, S3)
    
End Function

'@TestMethod
Public Sub SetIntersect()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    Setup.unpack S1, S2, S3
    
    Assert.AreEqual "SortedSet(1, 3, 5)", S1.Intersect(S1).ToString
    
    Assert.AreEqual "SortedSet()", S1.Intersect(S2).ToString
    Assert.AreEqual "SortedSet()", S2.Intersect(S1).ToString
    
    Assert.AreEqual "SortedSet(3, 5)", S3.Intersect(S1).ToString
    Assert.AreEqual "SortedSet(3, 5)", S1.Intersect(S3).ToString
    
End Sub

'@TestMethod
Public Sub SetUnion()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    Setup.unpack S1, S2, S3
    
    Assert.AreEqual "SortedSet(1, 3, 5)", S1.Union(S1).ToString
    
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", S1.Union(S2).ToString
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", S2.Union(S1).ToString
    
    Assert.AreEqual "SortedSet(1, 3, 5, 7)", S3.Union(S1).ToString
    Assert.AreEqual "SortedSet(1, 3, 5, 7)", S1.Union(S3).ToString
    
End Sub

'@TestMethod
Public Sub SetSymmetricDifference()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    Setup.unpack S1, S2, S3
    
    Assert.AreEqual "SortedSet()", S1.SymmetricDifference(S1).ToString
    
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", S1.SymmetricDifference(S2).ToString
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", S2.SymmetricDifference(S1).ToString
    
    Assert.AreEqual "SortedSet(1, 7)", S3.SymmetricDifference(S1).ToString
    Assert.AreEqual "SortedSet(1, 7)", S1.SymmetricDifference(S3).ToString
    
End Sub

'@TestMethod
Public Sub SetIsDisJoint()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    Setup.unpack S1, S2, S3
    
    Assert.IsFalse S1.IsDisJoint(S1)
    
    Assert.IsTrue S1.IsDisJoint(S2)
    Assert.IsTrue S2.IsDisJoint(S1)
    
    Assert.IsFalse S3.IsDisJoint(S1)
    Assert.IsFalse S1.IsDisJoint(S3)
    
End Sub
Private Function SubSetSetup() As Tuple

    Dim S1 As SortedSet
    Set S1 = SortedSet.Create(1, 3, 5)
    
    Dim S2 As SortedSet
    Set S2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5)
    
    Set SubSetSetup = Tuple.Pack(S1, S2, S3)
    
End Function
'@TestMethod
Public Sub SetIsSubSet()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    SubSetSetup.unpack S1, S2, S3
    
    Assert.IsTrue S1.IsSubSetOf(S1)
    
    Assert.IsFalse S1.IsSubSetOf(S2)
    Assert.IsFalse S2.IsSubSetOf(S1)
    
    Assert.IsTrue S3.IsSubSetOf(S1)
    Assert.IsFalse S1.IsSubSetOf(S3)
    
End Sub

'@TestMethod
Public Sub SetIsProperSubSet()

    Dim S1 As SortedSet, S2 As SortedSet, S3 As SortedSet
    SubSetSetup.unpack S1, S2, S3
    
    Assert.IsFalse S1.IsProperSubSetOf(S1)
    
    Assert.IsFalse S1.IsProperSubSetOf(S2)
    Assert.IsFalse S2.IsProperSubSetOf(S1)
    
    Assert.IsTrue S3.IsProperSubSetOf(S1)
    Assert.IsFalse S1.IsProperSubSetOf(S3)
    
End Sub
