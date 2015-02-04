Attribute VB_Name = "LambdaTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Option Explicit

Private Const SQUARE_LAMBDA As String = "(x) => x * x"
Private Const CONCAT_LAMBDA As String = "(a, b) => a & b"

'@TestMethod
Public Sub LambdaFromStringIsToString()
    
    Assert.AreEqual "Lambda[" & SQUARE_LAMBDA & "]", Lambda.FromString(SQUARE_LAMBDA).ToString, "Lambda[FromStr] = ToString"
        
End Sub
'@TestMethod
Public Sub LambdaFromStringExec()
    
    Assert.AreEqual 4, Lambda.FromString(SQUARE_LAMBDA)(2), "Lambda[(2) => 2 * 2 == 4]"
    
End Sub
'@TestMethod
Public Sub LambdaFromShortHand()

    Dim f As Lambda
    Set f = Lambda.FromShortHand("_ & _")
    
    Assert.AreEqual Lambda.FromString(CONCAT_LAMBDA).ToString, f.ToString
    Assert.AreEqual "Hi", f.Exec("H", "i")

End Sub
'@TestMethod
Public Sub LambdaCallExecTwice()
    
    Dim f As Lambda
    Set f = Lambda.FromString(SQUARE_LAMBDA)
    
    Assert.AreEqual 4, f(2), "lambda.exec x 1"
    Assert.AreEqual 9, f(3), "lambda.exec x 2"
    
End Sub
'@TestMethod
Public Sub LambdaResetsStaticVars()

    Dim f As Lambda
    Set f = Lambda.FromString(SQUARE_LAMBDA)
    
    Assert.AreEqual 1, StaticTest
    Assert.AreEqual 2, StaticTest
    f.Exec 1, 1
    Assert.AreEqual 3, StaticTest

End Sub
Private Function StaticTest() As Integer

    Static s As Integer
    s = s + 1
    
    StaticTest = s
    
End Function


