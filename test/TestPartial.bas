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
    Set keyPart = Partial.Make(itemGetter, Array(Null, "default value"))
    
    BatteryApplicable.Battery keyPart, "key", "default value"
    
    myDict.Append "key", "value"
    BatteryApplicable.Battery keyPart, "key", "value"

End Sub

'@TestMethod
Public Sub PartialPrependTest()

    Dim myDict As Dict
    Set myDict = Dict.Create
    
    Dim itemGetter As OnArgs
    Set itemGetter = OnArgs.Make("GetItem", VbMethod, myDict)
    
    Dim defaultPart As Partial
    Set defaultPart = Partial.Prepend(itemGetter, "key")
    
    BatteryApplicable.Battery defaultPart, "default value", "default value"
    
    myDict.Append "key", "value"
    BatteryApplicable.Battery defaultPart, "default value", "value"

End Sub
'@TestMethod
Public Sub PartialAppendTest()

    Dim myDict As Dict
    Set myDict = Dict.Create
    
    Dim itemGetter As OnArgs
    Set itemGetter = OnArgs.Make("GetItem", VbMethod, myDict)
    
    Dim keyPart As Partial
    Set keyPart = Partial.Append(itemGetter, "default Value")

    BatteryApplicable.Battery keyPart, "key", "default value"
    
    myDict.Append "key", "value"
    BatteryApplicable.Battery keyPart, "key", "value"

End Sub
