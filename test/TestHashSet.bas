Attribute VB_Name = "TestHashSet"
Option Explicit
Option Private Module
'@TestModule
Private Assert As New Rubberduck.AssertClass


'@TestMethod
Public Sub TestHashSetGist()
    
    Dim x As List
    Set x = List.Create(1, 2, 3)
    
    Dim y As List
    Set y = List.Create(1, 2, 3)
    
    Dim z As List
    Set z = y
    
    Dim hs As HashSet
    Set hs = HashSet.Create(x, x, y, y, z, z)
    
    Assert.areequal CLng(2), hs.Count
    Assert.IsTrue hs.Contains(x)
    Assert.IsTrue hs.Contains(y)
    Assert.IsTrue hs.Contains(z)
    
End Sub
'@TestMethod
Public Sub TestHashAndSortedEquals()

    Dim compatable As HashSet
    Set compatable = HashSet.Create(1, 2, 3)
    
    Dim incompatable As HashSet
    Set incompatable = HashSet.Create(List.Create(1, 2, 3))
    
    Dim ss As SortedSet
    Set ss = SortedSet.Create(1, 2, 3)

    On Error GoTo TestFail
    Assert.IsTrue Equals(compatable, ss)
    Assert.IsFalse Equals(incompatable, ss)
    On Error GoTo 0

Exit Sub
TestFail:
    Assert.Fail
    Resume Next
End Sub
'@TestMethod
Public Sub BatteryHashSet()

    Dim x As List
    Set x = List.Create(1, 2, 3)
    
    Dim y As List
    Set y = List.Create(1, 2, 3)
    
    Dim z As List
    Set z = y
    
    Dim otherSet As SortedSet
    Set otherSet = SortedSet.Create(1, 2, 3)
    
    Dim otherAssoc As Assoc
    Set otherAssoc = Assoc.Make(1, 2)
    
    Dim hsA As HashSet
    Set hsA = HashSet.Create(x, y, z)
    
    Dim hsB As HashSet
    Set hsB = HashSet.Create(x, otherSet)
    
    Dim hsC As HashSet
    Set hsC = HashSet.Create(otherAssoc, z)
    
    Dim super As HashSet
    Set super = HashSet.Create(x, y, z, otherAssoc, otherSet)
    
    Dim emptySet As HashSet
    Set emptySet = HashSet.Create()
    
    BatterySetLike.Battery hsA, hsB, hsC, emptySet, super

End Sub
