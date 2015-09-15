Attribute VB_Name = "Testcast"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'
' cast Test
' =========
'
'@TestMethod
Public Sub CastAssignPrimative()

    Dim x As Integer
    x = 1
    
    cast.Assign x, 2
    Assert.AreEqual 2, x
    
End Sub
'@TestMethod
Public Sub CastAssignArray()

    Dim a() As Variant
    a = List.Create(1, 2, 3).ToArray
    Assert.AreEqual "1 2 3", Join(a), "proving givens"
    
    Dim lower As Integer
    lower = LBound(a)
    
    cast.Assign a(lower), "x"
    Assert.AreEqual "x", a(lower)

End Sub
'@TestMethod
Public Sub CastAssignObject()

    Dim xs As List
    Set xs = List.Create(1, 2, 3)
    
    Dim ys As List
    Set ys = List.Create("A", "B", "C")
    
    ' Assign uses default "Set"
    cast.Assign xs, ys
    Assert.IsTrue xs.Equals(ys)
    Assert.AreEqual ObjPtr(ys), ObjPtr(xs)
    
    'Double Check
    ys.Append "D"
    Assert.IsTrue xs.Equals(ys)
    
End Sub
'
' CArray
' ------
'
'@TestMethod
Public Sub CastCArray()

End Sub


