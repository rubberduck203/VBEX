VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Test As VBAUnit.UnitTest

Friend Sub SetOutputStream(out As IOutput)
    Set Test = VBAUnit.UnitTestFactory.Create(TypeName(Me), out)
End Sub

Private Sub Class_Terminate()
    Test.Dispose
    Set Test = Nothing
End Sub


Public Sub CountShouldBe5()
    Test.Assert.Name = "CountShouldBe5"
    
    Dim e As Enumerable
    Set e = GetSortedEnumerable(5)
    
    Test.Assert.AreEqual 5, e.Count
    Test.Assert.IsFalse (6 = e.Count)
    
End Sub

Public Sub ShouldBeSorted()
   Test.Assert.Name = "ShouldBeSorted"
   
   Dim e As Enumerable
   Set e = GetSortedEnumerable(5)
   
   Test.Assert.IsTrue e.IsSorted

End Sub

Public Sub ShouldNotBeSorted()
    Test.Assert.Name = "ShouldNotBeSorted"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    Test.Assert.IsFalse e.IsSorted
End Sub

Public Sub MinShouldBe1()
    Test.Assert.Name = "MinShouldBe1"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    Test.Assert.AreEqual 1, e.Min
End Sub

Public Sub MaxShouldBe10()
    Test.Assert.Name = "MaxShouldBe10"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    Test.Assert.AreEqual 10, e.Max
End Sub

Private Function GetSortedEnumerable(numberOfItems As Long) As Enumerable
    Set GetSortedEnumerable = Enumerable.Range(1, numberOfItems)
End Function

Private Function GetUnsortedEnumerable() As Enumerable
    Dim e As Enumerable
    Set e = Enumerable.Create
    
    e.Add 4
    e.Add 1
    e.Add 2
    e.Add 10
    
    Set GetUnsortedEnumerable = e
End Function

