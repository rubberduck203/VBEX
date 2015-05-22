Attribute VB_Name = "BatteryApplicable"
Option Explicit
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
Public Sub test()

    Dim g, m, f, u
    Dim xs
    Set xs = List.Create(2, 3, 4, 5, 6, 7)
    Set g = Lambda.FromProper("(x) => x*x")
    Set m = Maybe.Some(xs)
    Set u = NameCall.OnArgs("Some", VbMethod, Maybe)
    Set f = NameCall.OnObject("Map", VbMethod, g)
    Dim h
    Set h = Composed.Make(u, f)
    
    console.PrintLine m.bind(h)
    
End Sub
