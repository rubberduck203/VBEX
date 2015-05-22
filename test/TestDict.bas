Attribute VB_Name = "TestDict"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

'
' Constructors
' ------------
'
Public Sub DictTestEmpty(ByVal d As Dict)
    Assert.IsNotNothing d, "Empty Dict is not nothing"
    Assert.AreEqual CLng(0), d.Count, "Empty Dict count = 3"
    Assert.AreEqual "Dict()", d.Show, "Empty Dict Show"
End Sub

Public Sub DictTestNonEmpty(ByVal d As Dict)
    Assert.IsNotNothing d, "NonEmpty Dict is not nothing"
    Assert.AreEqual CLng(3), d.Count, "NonEmpty Dict count = 3"
    Assert.AreEqual "Dict(1 -> 2, 3 -> 4, 5 -> 6)", d.Show, "NonEmpty Dict Show"
End Sub

'@TestMethod
Public Sub DictEmptyFromAssocs()
    DictTestEmpty Dict.FromAssocs(List.Create())
End Sub

'@TestMethod
Public Sub DictNonEmptyFromAssocs()
    DictTestNonEmpty Dict.FromAssocs(List.Create(Assoc.Make(1, 2), Assoc.Make(3, 4), Assoc.Make(5, 6)))
End Sub

'@TestMethod
Public Sub DictCreateEmptyDict()
    DictTestEmpty Dict.Create()
End Sub

'@TestMethod
Public Sub DictCreateNonEmptyDict()
    DictTestNonEmpty Dict.Create(Assoc.Make(1, 2), Assoc.Make(3, 4), Assoc.Make(5, 6))
End Sub

'@TestMethod
Public Sub DictCopyEmptyDict()
    DictTestEmpty Dict.Copy(Dict.Create())
End Sub

'@TestMethod
Public Sub DictCopyNonEmptyDict()
    DictTestNonEmpty Dict.Copy(Dict.Create(Assoc.Make(1, 2), Assoc.Make(3, 4), Assoc.Make(5, 6)))
End Sub

'@TestMethod
Public Sub DictCopyIsCopy()
    Dim orig As Dict, cpy As Dict
    Set orig = Dict.Create
    Set cpy = Dict.Copy(orig)
    Assert.AreNotEqual ObjPtr(orig), ObjPtr(cpy), "Copy is a new instance"
End Sub

'@TestMethod
Public Sub DictEmptyFromLists()
    DictTestEmpty Dict.FromLists(List.Create(), List.Create())
End Sub

'@TestMethod
Public Sub DictNonEmptyFromLists()
    DictTestNonEmpty Dict.FromLists(List.Create(1, 3, 5), List.Create(2, 4, 6))
End Sub


'@TestMethod
Public Sub DictKeysAndValues()
    
    Dim ks As List
    Set ks = List.Create(1, 2, 3)
    
    Dim vs As List
    Set vs = List.Create("a", "b", "c")
    
    Dim d As Dict
    Set d = Dict.FromLists(ks, vs)
    
    Assert.IsTrue ks.Equals(d.Keys)
    Assert.IsTrue vs.Equals(d.Values)
    
End Sub
