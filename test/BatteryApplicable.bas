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

    Assert.AreEqual y, f.Apply(x)

End Sub
Private Sub TestPartial(ByVal f As IApplicable, ByVal x, ByVal y)

    Assert.AreEqual y, f.Partial(x).Apply()
    Assert.AreEqual y, f.Partial(Empty).Apply(x)
    Assert.AreEqual y, f.AsPartial(xArray(x)).Apply()
    Assert.AreEqual y, f.AsPartial(xArray(Empty)).Apply(x)

End Sub
Private Sub TestDelay(ByVal f As IApplicable, ByVal x, ByVal y)

    Assert.AreEqual y, f.Delay(x).Evaluate()
    Assert.AreEqual y, f.AsDelay(xArray(x)).Evaluate()

End Sub
