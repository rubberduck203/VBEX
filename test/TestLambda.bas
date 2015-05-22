Attribute VB_Name = "TestLambda"
'@TestModule
Option Explicit
Private Assert As New Rubberduck.AssertClass



'@TestMethod
Public Sub LambdaFromProper()
    
    Dim f As Lambda
    Set f = Lambda.FromProper("(x) => x * x")
    
    BatteryApplicable.Battery f, 2, 4
    
End Sub
'@TestMethod
Public Sub LambdaFromShort()

    Dim f As Lambda
    Set f = Lambda.FromShort("_ + 13")
    
    BatteryApplicable.Battery f, 11, 24
     
End Sub

