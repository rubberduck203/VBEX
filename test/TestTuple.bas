Attribute VB_Name = "TestTuple"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

Private Sub TupleTestEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "Empty Tuple is not nothing"
    Assert.AreEqual CLng(0), t.Count, "Empty Tuple count"
    Assert.AreEqual "Tuple()", t.Show, "Empty Tuple Show"
End Sub
Private Sub TupleTestNonEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "NonEmpty Tuple is not nothing"
    Assert.AreEqual CLng(3), t.Count, "NonEmpty Tuple count"
    Assert.AreEqual "Tuple(1, 2, 3)", t.Show, "NonEmpty Tuple Show"
End Sub

'@TestMethod
Public Sub TuplePackEmpty()
    TupleTestEmpty Tuple.Pack()
End Sub

'@TestMethod
Public Sub TuplePackNonEmpty()
    TupleTestNonEmpty Tuple.Pack(1, 2, 3)
End Sub

'@TestMethod
Public Sub TupleImpodeEmpty()
    TupleTestEmpty Tuple.Implode(Array())
End Sub

'@TestMethod
Public Sub TupleImpodeNonEmpty()
   TupleTestNonEmpty Tuple.Implode(Array(1, 2, 3))
End Sub

'@TestMethod
Public Sub TupleImpodeNonEmptyOffset()

    'TODO: Shouldn't need to be variant
    Dim a(1 To 3) As Variant
    a(1) = 1
    a(2) = 2
    a(3) = 3
    
    TupleTestNonEmpty Tuple.Implode(a)
    
End Sub

'@TestMethod
Public Sub TupleUnpackEmpty()
    'TODO: no error
    Dim t As Tuple
    Set t = Tuple.Pack()

    t.unpack
    
End Sub
Private Function Setup() As Tuple

   Set Setup = Tuple.Pack(1, 2, 3)
   
End Function
'@TestMethod
Public Sub TupleUnpackNonEmpty()

    Dim x As Integer, y As Integer, z As Integer
    Setup.unpack x, y, z
    
    Assert.AreEqual "1 2 3", Join(Array(x, y, z))
    
End Sub

'@TestMethod
Public Sub TupleExplodeNonEmpty()

    Dim a(0 To 2) As Variant
    Setup.Explode a
    
    Assert.AreEqual "1 2 3", Join(a)
    
End Sub

'@TestMethod
Public Sub TupleExplodeNonEmptyOffset()

    Dim a(1 To 3) As Variant
    Setup.Explode a
    
    Assert.AreEqual "1 2 3", Join(a)
    
End Sub
