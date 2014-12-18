Attribute VB_Name = "DictTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass

'
' Constructors
' ------------
'
Public Sub TestEmpty(ByVal d As Dict)
    Assert.IsNotNothing d, "Empty Dict is not nothing"
    Assert.AreEqual CLng(0), d.Count, "Empty Dict count = 3"
    Assert.AreEqual "{}", d.ToString, "Empty Dict ToString"
End Sub

Public Sub TestNonEmpty(ByVal d As Dict)
    Assert.IsNotNothing d, "NonEmpty Dict is not nothing"
    Assert.AreEqual CLng(3), d.Count, "NonEmpty Dict count = 3"
    Assert.AreEqual "{1: 2, 3: 4, 5: 6}", d.ToString, "NonEmpty Dict ToString"
End Sub

'@TestMethod
Public Sub EmptyFromTuples()
    TestEmpty Dict.FromTuples(List.Create())
End Sub

'@TestMethod
Public Sub NonEmptyFromTuples()
    TestNonEmpty Dict.FromTuples(List.Create(Tuple.pack(1, 2), Tuple.pack(3, 4), Tuple.pack(5, 6)))
End Sub

'@TestMethod
Public Sub CreateEmptyDict()
    TestEmpty Dict.Create()
End Sub

'@TestMethod
Public Sub CreateNonEmptyDict()
    TestNonEmpty Dict.Create(Tuple.pack(1, 2), Tuple.pack(3, 4), Tuple.pack(5, 6))
End Sub

'@TestMethod
Public Sub CopyEmptyDict()
    TestEmpty Dict.Copy(Dict.Create())
End Sub

'@TestMethod
Public Sub CopyNonEmptyDict()
    TestNonEmpty Dict.Copy(Dict.Create(Tuple.pack(1, 2), Tuple.pack(3, 4), Tuple.pack(5, 6)))
End Sub

'@TestMethod
Public Sub CopyIsCopy()
    Dim orig As Dict, cpy As Dict
    Set orig = Dict.Create
    Set cpy = Dict.Copy(orig)
    Assert.AreNotEqual ObjPtr(orig), ObjPtr(cpy), "Copy is a new instance"
End Sub

'@TestMethod
Public Sub EmptyFromLists()
    TestEmpty Dict.FromLists(List.Create(), List.Create())
End Sub

'@TestMethod
Public Sub NonEmptyFromLists()
    TestNonEmpty Dict.FromLists(List.Create(1, 3, 5), List.Create(2, 4, 6))
End Sub

