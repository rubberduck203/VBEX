Attribute VB_Name = "BatteryMonadic"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
' Monadic Battery
' ===============
'
Public Sub Battery(ByVal monad As IMonadic)

    Dim x
    x = 2
    
    Dim f As IApplicable
    Set f = Lambda.FromShort("_ * 2")
    
    Dim g As IApplicable
    Set g = Lambda.FromShort("_ + 13")
    
    Dim u As IApplicable
    Set u = ApplyUnit(monad)
    
    Dim uf As IApplicable
    Set uf = Composed.Make(u, f)
    
    Dim ug As IApplicable
    Set ug = Composed.Make(u, g)

    Dim m As IMonadic
    Set m = u.Apply(x)
    
    Associativity m, uf, ug
    LeftUnit m, x, uf
    RightUnit m

End Sub
'
' Tests
' -----
'
' TODO: Closure
'
Private Sub Associativity(ByVal m As IMonadic, ByVal f As IApplicable, _
        ByVal g As IApplicable)

    ' monad.Bind(f).Bind(g) == monad.Bind(x => f(x).Bind(g))
    Dim leftSide As IEquatable
    Set leftSide = m.bind(f).bind(g)

    Dim nc As IApplicable ' .Bind(g)
    Set nc = OnObject.Make("Bind", VbMethod, g)
    
    Dim h As IApplicable ' nc.Apply(f(x)) ==  f(x).Bind(g)
    Set h = Composed.Make(nc, f)

    Dim rightSide As IEquatable
    Set rightSide = m.bind(h)

    Dim result As Boolean
    result = Equals(leftSide, rightSide)
    Assert.IsTrue result

End Sub
' I don't think I have this correct
Private Sub LeftUnit(ByVal m As IMonadic, ByVal x, ByVal f As IApplicable)
    ' unit(x).Bind(f) == f(x)

    Dim leftSide As IEquatable
    Set leftSide = m.Unit(x).bind(f)
    
    Dim rightSide As IEquatable
    Set rightSide = f.Apply(x)
    
    Dim result As Boolean
    result = Equals(leftSide, rightSide)
    Assert.IsTrue result

End Sub
Private Sub RightUnit(ByVal m As IMonadic)
    ' m.Bind(unit) = m
    
    Dim u As IApplicable
    Set u = OnArgs.Make("Unit", VbMethod, m)
    
    Dim leftSide As IEquatable
    Set leftSide = m.bind(u)
    
    Dim result As Boolean
    result = Equals(leftSide, m)
    Assert.IsTrue result
    
End Sub
