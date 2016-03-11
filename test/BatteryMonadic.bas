Attribute VB_Name = "BatteryMonadic"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
' Monadic Battery
' ===============
'
Public Sub Battery(ByVal monad As Monadic)

    Dim x
    x = 2
    
    Dim f As Applicable
    Set f = Lambda.FromShort("_ * 2")
    
    Dim g As Applicable
    Set g = Lambda.FromShort("_ + 13")
    
    Dim u As Applicable
    Set u = ApplyUnit(monad)
    
    Dim uf As Applicable
    Set uf = Composed.Make(u, f)
    
    Dim ug As Applicable
    Set ug = Composed.Make(u, g)

    Dim m As Monadic
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
Private Sub Associativity(ByVal m As Monadic, ByVal f As Applicable, _
        ByVal g As Applicable)

    ' monad.Bind(f).Bind(g) == monad.Bind(x => f(x).Bind(g))
    Dim leftSide As Equatable
    Set leftSide = m.Bind(f).Bind(g)

    Dim nc As Applicable ' .Bind(g)
    Set nc = OnObject.Create("Bind", VbMethod, g)
    
    Dim h As Applicable ' nc.Apply(f(x)) ==  f(x).Bind(g)
    Set h = Composed.Make(nc, f)

    Dim rightSide As Equatable
    Set rightSide = m.Bind(h)

    Dim result As Boolean
    result = Equals(leftSide, rightSide)
    Assert.IsTrue result

End Sub
' I don't think I have this correct
Private Sub LeftUnit(ByVal m As Monadic, ByVal x, ByVal f As Applicable)
    ' unit(x).Bind(f) == f(x)

    Dim leftSide As Equatable
    Set leftSide = m.Unit(x).Bind(f)
    
    Dim rightSide As Equatable
    Set rightSide = f.Apply(x)
    
    Dim result As Boolean
    result = Equals(leftSide, rightSide)
    Assert.IsTrue result

End Sub
Private Sub RightUnit(ByVal m As Monadic)
    ' m.Bind(unit) = m
    
    Dim u As Applicable
    Set u = OnArgs.Make("Unit", VbMethod, m)
    
    Dim leftSide As Equatable
    Set leftSide = m.Bind(u)
    
    Dim result As Boolean
    result = Equals(leftSide, m)
    Assert.IsTrue result
    
End Sub
