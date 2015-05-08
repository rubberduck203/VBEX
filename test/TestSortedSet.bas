Attribute VB_Name = "TestSortedSet"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

Private Sub SetEmptyConstructorTest(ByVal emptyset As SortedSet)
    
    Assert.IsNotNothing emptyset, "Empty set is not nothing"
    Assert.AreEqual "SortedSet()", emptyset.Show, "Emptyset repr SortedSet()"
    Assert.AreEqual CLng(0), emptyset.Count, "emptyset count = 0"
    
End Sub
Private Sub SetNonEmptyConstructorTest(ByVal nonempty As SortedSet)

    Assert.IsNotNothing nonempty, "nonempty set is not nothing"
    Assert.AreEqual "SortedSet(1, 2, 3)", nonempty.Show, "nonempty repr SortedSet(1, 2, 3)"
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

    Dim s1 As SortedSet
    Set s1 = SortedSet.Create(1, 3, 5)
    
    Dim s2 As SortedSet
    Set s2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5, 7)
    
    Assert.AreEqual "SortedSet()", s1.Difference(s1).Show
    Assert.AreEqual "SortedSet(1, 3, 5)", s1.Difference(s2).Show
    Assert.AreEqual "SortedSet(2, 4, 6)", s2.Difference(s1).Show
    
    Assert.AreEqual "SortedSet(7)", S3.Difference(s1).Show
    Assert.AreEqual "SortedSet(1)", s1.Difference(S3).Show
    
End Sub

Private Function Setup() As Tuple
    
    Dim s1 As SortedSet
    Set s1 = SortedSet.Create(1, 3, 5)
    
    Dim s2 As SortedSet
    Set s2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5, 7)
    
    Set Setup = Tuple.Pack(s1, s2, S3)
    
End Function

'@TestMethod
Public Sub SetIntersect()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    Setup.unpack s1, s2, S3
    
    Assert.AreEqual "SortedSet(1, 3, 5)", s1.Intersect(s1).Show
    
    Assert.AreEqual "SortedSet()", s1.Intersect(s2).Show
    Assert.AreEqual "SortedSet()", s2.Intersect(s1).Show
    
    Assert.AreEqual "SortedSet(3, 5)", S3.Intersect(s1).Show
    Assert.AreEqual "SortedSet(3, 5)", s1.Intersect(S3).Show
    
End Sub

'@TestMethod
Public Sub SetUnion()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    Setup.unpack s1, s2, S3
    
    Assert.AreEqual "SortedSet(1, 3, 5)", s1.Union(s1).Show
    
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", s1.Union(s2).Show
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", s2.Union(s1).Show
    
    Assert.AreEqual "SortedSet(1, 3, 5, 7)", S3.Union(s1).Show
    Assert.AreEqual "SortedSet(1, 3, 5, 7)", s1.Union(S3).Show
    
End Sub

'@TestMethod
Public Sub SetSymmetricDifference()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    Setup.unpack s1, s2, S3
    
    Assert.AreEqual "SortedSet()", s1.SymmetricDifference(s1).Show
    
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", s1.SymmetricDifference(s2).Show
    Assert.AreEqual "SortedSet(1, 2, 3, 4, 5, 6)", s2.SymmetricDifference(s1).Show
    
    Assert.AreEqual "SortedSet(1, 7)", S3.SymmetricDifference(s1).Show
    Assert.AreEqual "SortedSet(1, 7)", s1.SymmetricDifference(S3).Show
    
End Sub

'@TestMethod
Public Sub SetIsDisJoint()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    Setup.unpack s1, s2, S3
    
    Assert.IsFalse s1.IsDisJoint(s1)
    
    Assert.IsTrue s1.IsDisJoint(s2)
    Assert.IsTrue s2.IsDisJoint(s1)
    
    Assert.IsFalse S3.IsDisJoint(s1)
    Assert.IsFalse s1.IsDisJoint(S3)
    
End Sub
Private Function SubSetSetup() As Tuple

    Dim s1 As SortedSet
    Set s1 = SortedSet.Create(1, 3, 5)
    
    Dim s2 As SortedSet
    Set s2 = SortedSet.Create(2, 4, 6)
    
    Dim S3 As SortedSet
    Set S3 = SortedSet.Create(3, 5)
    
    Set SubSetSetup = Tuple.Pack(s1, s2, S3)
    
End Function
'@TestMethod
Public Sub SetIsSubSet()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    SubSetSetup.unpack s1, s2, S3
    
    Assert.IsTrue s1.IsSubSetOf(s1)
    
    Assert.IsFalse s1.IsSubSetOf(s2)
    Assert.IsFalse s2.IsSubSetOf(s1)
    
    Assert.IsTrue S3.IsSubSetOf(s1)
    Assert.IsFalse s1.IsSubSetOf(S3)
    
End Sub

'@TestMethod
Public Sub SetIsProperSubSet()

    Dim s1 As SortedSet, s2 As SortedSet, S3 As SortedSet
    SubSetSetup.unpack s1, s2, S3
    
    Assert.IsFalse s1.IsProperSubSetOf(s1)
    
    Assert.IsFalse s1.IsProperSubSetOf(s2)
    Assert.IsFalse s2.IsProperSubSetOf(s1)
    
    Assert.IsTrue S3.IsProperSubSetOf(s1)
    Assert.IsFalse s1.IsProperSubSetOf(S3)
    
End Sub
