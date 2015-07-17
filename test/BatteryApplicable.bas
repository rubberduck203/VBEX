Attribute VB_Name = "BatteryApplicable"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
' Applicable Battery
' ==================
'
Public Sub Battery(ByVal f As IApplicable, ByVal x, ByVal y)

    TestApply f, x, y

End Sub
'
' Tests
' -----
'
Private Sub TestApply(ByVal f As IApplicable, ByVal x, ByVal y)

    Assert.areequal y, f.Apply(x)

End Sub
