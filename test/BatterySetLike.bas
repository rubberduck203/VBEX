Attribute VB_Name = "BatterySetLike"
Option Explicit
Private Assert As New Rubberduck.AssertClass

Public Sub Battery(ByVal setA As ISetLike, _
        ByVal setB As ISetLike, _
        ByVal setC As ISetLike, _
        ByVal emptySet As ISetLike
        ByVal super As ISetLike)

    IdentityLaw setA, super, emptySet
    DomainLaw setA, super, emptySet
    IdempotentLaw setA
    CommutativeLaw setA, setB
    AssociatveLaw setA, setB, setC
    DistributiveLaw setA, setB, setC

End Sub

Private Sub IdentityLaw(ByVal setA As ISetLike,ByVal  super As ISetLike, ByVal emptySet As ISetLike)

    ' A U 0 = A
    Assert.IsTrue Equals(setA.Union(emptySet), setA)
    'A n U = A
    Assert.IsTrue Equals(setA.Intersect(super), setA)
    
End Sub
Private Sub DomainLaw(ByVal setA As ISetLike, ByVal super As ISetLike, ByVal emptySet As ISetLike)

    ' A u U = U
    Assert.IsTrue Equals(setA.Union(super), super)
    ' A n 0 = 0
    Assert.IsTrue Equals(setA.Intersect(emptySet), emptySet)

End Sub
Private Sub IdempotentLaw(ByVal setA As ISetLike)

    ' A u A = A
    Assert.IsTrue Equals(setA.Union(setA), setA)
    ' A n A = A
    Assert.IsTrue Equals(setA.Intersect(setA), setA)

End Sub
Private Sub CommutativeLaw(ByVal setA As ISetLike, ByVal setB As ISetLike)

    ' A u B = B u A
    Assert.IsTrue Equals(setA.Union(setB), setB.Union(setA))
    ' A n B = B n A
    Assert.IsTrue Equals(setA.Intersect(setB), setB.Intersect(setA))

End Sub
Private Sub AssociatveLaw(ByVal setA As ISetLike, ByVal setB As ISetLike, ByVal setC As ISetLike)

    ' (A u B) u C = A u (B u C)
    Dim lhsLaw1 As IEquatable
    Set lhsLaw1 = setA.Union(setB).Union(setC)

    Dim rhsLaw1 As IEquatable
    Set rhsLaw1 = setA.Union(setB.Union(setC))

    Assert.IsTrue Equals(lhsLaw1, rhsLaw1) 

    ' (A n B) n C = A n (B n C)
    Dim lhsLaw2 As IEquatable
    Set lhsLaw2 = setA.Intersect(setB).Intersect(setC)

    Dim rhsLaw2 As IEquatable
    Set rhsLaw2 = setA.Intersect(setB.Intersect(setC))

    Assert.IsTrue Equals(lhsLaw2, rhsLaw2) 

End Sub
Private Sub DistributiveLaw(ByVal setA As ISetLike, ByVal setB As ISetLike, ByVal setC As ISetLike)

    ' A u (B n C) = (A u B) n (A u C)
    Dim lhsLaw1 As IEquatable
    Set lhsLaw1 = setA.Union(setB.Intersect(setC))

    Dim rhsLaw1 As IEquatable
    Set rhsLaw1 = (setA.Union(setB)).Intersect(setA.Union(setC))

    Assert.IsTrue Equals(lhsLaw1, rhsLaw1) 


    ' A n (B u C) = (A n B) u (A n C)
    Dim lhsLaw2 As IEquatable
    Set lhsLaw2 = setA.Intersect(setB.Uniont(setC))

    Dim rhsLaw2 As IEquatable
    Set rhsLaw2 = (setA.Intersect(setB)).Union(setA.Intersect(setC))

    Assert.IsTrue Equals(lhsLaw2, rhsLaw2) 

End Sub
