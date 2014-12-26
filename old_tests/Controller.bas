Attribute VB_Name = "Controller"
Option Explicit

Public Sub RunEnumerableTests()
    Dim Test As New TestEnumerable
    Test.SetOutputStream VBAUnit.Console
    
    Test.CountShouldBe5
    Test.ShouldBeSorted
    Test.ShouldNotBeSorted
    Test.MinShouldBe1
    Test.MaxShouldBe10

End Sub

Public Sub RunVbeCodeModuleTests()
    Dim Test As New TestVbeCodeModule_IsSignature
    Test.SetOutputStream VBAUnit.Console
    
    Test.CommentedOutLineShouldBeFalse
End Sub
