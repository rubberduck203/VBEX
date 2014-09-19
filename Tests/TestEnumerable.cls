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

Private test As VBAUnit.UnitTest

Friend Sub SetOutputStream(out As IOutput)
    Set test = VBAUnit.UnitTestFactory.Create(TypeName(Me), out)
End Sub

Private Sub Class_Terminate()
    test.Dispose
    Set test = Nothing
End Sub


Public Sub CountShouldBe5()
    test.Assert.Name = "CountShouldBe5"
    
    Dim e As Enumerable
    Set e = GetSortedEnumerable(5)
    
    test.Assert.AreEqual 5, e.Count
    test.Assert.IsFalse (6 = e.Count)
    
End Sub

Public Sub ShouldBeSorted()
   test.Assert.Name = "ShouldBeSorted"
   
   Dim e As Enumerable
   Set e = GetSortedEnumerable(5)
   
   test.Assert.IsTrue e.IsSorted

End Sub

Public Sub ShouldNotBeSorted()
    test.Assert.Name = "ShouldNotBeSorted"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    test.Assert.IsFalse e.IsSorted
End Sub

Public Sub MinShouldBe1()
    test.Assert.Name = "MinShouldBe1"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    test.Assert.AreEqual 1, e.Min
End Sub

Public Sub MaxShouldBe10()
    test.Assert.Name = "MaxShouldBe10"
    
    Dim e As Enumerable
    Set e = GetUnsortedEnumerable
    
    test.Assert.AreEqual 10, e.Max
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
