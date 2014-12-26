Attribute VB_Name = "TupleTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass

Public Sub TupleTestEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "Empty Tuple is not nothing"
    Assert.AreEqual CLng(0), t.Count, "Empty Tuple count"
    Assert.AreEqual "()", t.ToString, "Empty Tuple tostring"
End Sub
Public Sub TupleTestNonEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "NonEmpty Tuple is not nothing"
    Assert.AreEqual CLng(3), t.Count, "NonEmpty Tuple count"
    Assert.AreEqual "(1, 2, 3)", t.ToString, "NonEmpty Tuple tostring"
End Sub
'@TestMethod
Public Sub TuplePackEmptyTuple()
    TupleTestEmpty Tuple.pack()
End Sub
'@TestMethod
Public Sub TuplePackNonEmptyTuple()
    TupleTestNonEmpty Tuple.pack(1, 2, 3)
End Sub
'@TestMethod
Public Sub TupleImpodeEmptyTuple()
    TupleTestEmpty Tuple.Implode(Array())
End Sub
'@TestMethod
Public Sub TupleImpodeNonEmptyTuple()
   TupleTestNonEmpty Tuple.Implode(Array(1, 2, 3))
End Sub
