Attribute VB_Name = "TestMaybe"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub MaybeMonadicTest()

    Dim m As Maybe
    Set m = Maybe.Some(2)
    
    BatteryMonadic.Battery m

End Sub
'@TestMethod
Public Sub MaybeCanDefaultWithApplicationRun()

    Dim m As Maybe
    Set m = Maybe.Some("C:\Some\Path\yo.txt")
    
    Dim baser As InternalDelegate
    Set baser = InternalDelegate.Make("path.BaseName")
    
    On Error GoTo typeFailed
    Assert.AreEqual "yo.txt", baser.Apply(m)

Exit Sub
typeFailed:
    Assert.AreEqual CLng(13), Err.Number, "Error Number is not 13"
    Assert.Fail "Maybe did not default when called by application run"
End Sub
