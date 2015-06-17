Attribute VB_Name = "TestByName"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass


'@TestMethod
Public Sub ByNameCreateTest()

    Dim op As OnArgs
    Set op = OnArgs.Make("GetItem", VbMethod, Dict.Create(Assoc.Make("key", "value")))
    
    Dim bn As ByName
    Set bn = ByName.Create(op, "key", "default")
    Assert.AreEqual bn.Evaluate, "value"

End Sub
'@TestMethod
Public Sub ByNameMakeTest()

    Dim op As OnArgs
    Set op = OnArgs.Make("GetItem", VbMethod, Dict.Create(Assoc.Make("key", "value")))
    
    Dim args As Tuple
    Set args = Tuple.Pack("none", "default")
    
    Dim bn As ByName
    Set bn = ByName.Make(op, args)
    Assert.AreEqual bn.Evaluate, "default"

End Sub
