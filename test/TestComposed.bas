Attribute VB_Name = "TestComposed"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub ComposeTest()

    Dim innerOp As Lambda
    Set innerOp = Lambda.FromShort("_ + 12")
    
    Dim outerOp As Lambda
    Set outerOp = Lambda.FromShort("_ * 10")
    
    Dim comp As Composed
    Set comp = Composed.Make(outerOp, innerOp)
    
    BatteryApplicable.Battery comp, 5, 170

End Sub
'@TestMethod
Public Sub ComposeRecurseTest()

    Dim firstOp As Lambda
    Set firstOp = Lambda.FromShort("_ + 12")
    
    Dim secondOp As Lambda
    Set secondOp = Lambda.FromShort("_ * 10")
    
    Dim comp As Composed
    Set comp = Composed.Make(secondOp, firstOp)
    
    Dim thirdOp As Lambda
    Set thirdOp = Lambda.FromShort("_ \ 2")

    BatteryApplicable.Battery Composed.Make(thirdOp, comp), 5, 85

End Sub
