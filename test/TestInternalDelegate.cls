Attribute VB_Name = "TestInternalDelegate"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub TestPass()

    Dim idg As InternalDelegate
    Set idg = InternalDelegate.Make("MaxValue")

    Dim arg 
    arg = Array(1, 2, 4, 2, 100, 2, 3, 20, 3)

    Dim result As Integer
    result = 100
    
    BatteryIApplicable.Battery idg, arg, result
    
End Sub
