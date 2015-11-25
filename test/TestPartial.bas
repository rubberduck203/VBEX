Attribute VB_Name = "TestPartial"
Option Explicit

'@TestModule
Option Private Module
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub PartialMakeTest()

    Dim myDict As Dict
    Set myDict = Dict.Create
    
    Dim itemGetter As OnArgs
    Set itemGetter = OnArgs.Make("GetItem", VbMethod, myDict)
    
    Dim keyPart As Partial
    Set keyPart = Partial.Make(itemGetter, Array(Empty, "default value"))
    
    BatteryApplicable.Battery keyPart, "key", "default value"
    
    myDict.Append "key", "value"
    BatteryApplicable.Battery keyPart, "key", "value"

End Sub
