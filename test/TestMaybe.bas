Attribute VB_Name = "TestMaybe"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub MaybeMonadicTest()

    Dim m As Maybe
    Set m = Maybe.Some(2)
    
    BatteryMonadic.Battery m

End Sub

