Attribute VB_Name = "DictTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

'
' Constructors
' ------------
'
Public Sub DictTestEmpty(ByVal d As dict)
    Assert.IsNotNothing d, "Empty Dict is not nothing"
    Assert.AreEqual CLng(0), d.Count, "Empty Dict count = 3"
    Assert.AreEqual "{}", d.ToString, "Empty Dict ToString"
End Sub

Public Sub DictTestNonEmpty(ByVal d As dict)
    Assert.IsNotNothing d, "NonEmpty Dict is not nothing"
    Assert.AreEqual CLng(3), d.Count, "NonEmpty Dict count = 3"
    Assert.AreEqual "{1: 2, 3: 4, 5: 6}", d.ToString, "NonEmpty Dict ToString"
End Sub

'@TestMethod
Public Sub DictEmptyFromTuples()
    DictTestEmpty dict.FromTuples(List.Create())
End Sub

'@TestMethod
Public Sub DictNonEmptyFromTuples()
    DictTestNonEmpty dict.FromTuples(List.Create(Tuple.Pack(1, 2), Tuple.Pack(3, 4), Tuple.Pack(5, 6)))
End Sub

'@TestMethod
Public Sub DictCreateEmptyDict()
    DictTestEmpty dict.Create()
End Sub

'@TestMethod
Public Sub DictCreateNonEmptyDict()
    DictTestNonEmpty dict.Create(Tuple.Pack(1, 2), Tuple.Pack(3, 4), Tuple.Pack(5, 6))
End Sub

'@TestMethod
Public Sub DictCopyEmptyDict()
    DictTestEmpty dict.Copy(dict.Create())
End Sub

'@TestMethod
Public Sub DictCopyNonEmptyDict()
    DictTestNonEmpty dict.Copy(dict.Create(Tuple.Pack(1, 2), Tuple.Pack(3, 4), Tuple.Pack(5, 6)))
End Sub

'@TestMethod
Public Sub DictCopyIsCopy()
    Dim orig As dict, cpy As dict
    Set orig = dict.Create
    Set cpy = dict.Copy(orig)
    Assert.AreNotEqual ObjPtr(orig), ObjPtr(cpy), "Copy is a new instance"
End Sub

'@TestMethod
Public Sub DictEmptyFromLists()
    DictTestEmpty dict.FromLists(List.Create(), List.Create())
End Sub

'@TestMethod
Public Sub DictNonEmptyFromLists()
    DictTestNonEmpty dict.FromLists(List.Create(1, 3, 5), List.Create(2, 4, 6))
End Sub


'@TestMethod
Public Sub DictKeysAndValues()
    
    Dim ks As List
    Set ks = List.Create(1, 2, 3)
    
    Dim vs As List
    Set vs = List.Create("a", "b", "c")
    
    Dim d As dict
    Set d = dict.FromLists(ks, vs)
    
    Assert.IsTrue ks.Equals(d.Keys)
    Assert.IsTrue vs.Equals(d.Values)
    
End Sub
