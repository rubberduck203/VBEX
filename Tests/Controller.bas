Attribute VB_Name = "Controller"
Option Explicit

Public Sub RunEnumerableTests()
    Dim test As New TestEnumerable
    test.SetOutputStream VBAUnit.Console
    
    test.CountShouldBe5
    test.ShouldBeSorted
    test.ShouldNotBeSorted
    test.MinShouldBe1
    test.MaxShouldBe10

End Sub
