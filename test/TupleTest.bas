Attribute VB_Name = "TupleTest"
'@TestModule
Private Assert As New Rubberduck.AssertClass

Public Sub TestEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "Empty Tuple is not nothing"
    Assert.AreEqual CLng(0), t.Count, "Empty Tuple count"
    Assert.AreEqual "()", t.ToString, "Empty Tuple tostring"
End Sub
Public Sub TestNonEmpty(ByVal t As Tuple)
    Assert.IsNotNothing t, "NonEmpty Tuple is not nothing"
    Assert.AreEqual CLng(3), t.Count, "NonEmpty Tuple count"
    Assert.AreEqual "(1, 2, 3)", t.ToString, "NonEmpty Tuple tostring"
End Sub
'@TestMethod
Public Sub PackEmptyTuple()
    TestEmpty Tuple.pack()
End Sub
'@TestMethod
Public Sub PackNonEmptyTuple()
    TestNonEmpty Tuple.pack(1, 2, 3)
End Sub
'@TestMethod
Public Sub ImpodeEmptyTuple()
    TestEmpty Tuple.Implode(Array())
End Sub
'@TestMethod
Public Sub ImpodeNonEmptyTuple()
   TestNonEmpty Tuple.Implode(Array(1, 2, 3))
End Sub
