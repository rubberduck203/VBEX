Attribute VB_Name = "TestNameCall"
'@TestModule
Option Explicit
Option Private Module

'@TestMethod
Public Sub OnArgsGetTest()

    Dim xs As List
    Set xs = List.Create("a", "b")

    Dim nc As NameCall
    Set nc = NameCall.OnArgs("Item", VbGet, xs)

    BatteryApplicable.Battery nc, 2, "b"

End Sub
'@TestMethod
Public Sub OnArgsMethodTest()

    Dim xs As SortedSet
    Set xs = SortedSet.Create(1, 2, 3)

    Dim nc As NameCall
    Set nc = NameCall.OnArgs("Contains", VbMethod, xs)

    BatteryApplicable.Battery nc, 2, True
    BatteryApplicable.Battery nc, 4, False

End Sub
'@TestMethod
Public Sub OnObjectTest()

    Dim s1 As SortedSet
    Set s1 = SortedSet.Create(1, 2, 3)
    
    Dim s2 As SortedSet
    Set s2 = SortedSet.Create("a", "b", "c")

    Dim nc As NameCall
    Set nc = NameCall.OnObject("Contains", VbMethod, 2)

    BatteryApplicable.Battery nc, s1, True
    BatteryApplicable.Battery nc, s2, False

End Sub
''@TestMethod
'Public Sub OnMethodTest()
'
'    Dim s1 As SortedSet
'    Set s1 = SortedSet.Create(1, 2, 3)
'
'    Dim s2 As SortedSet
'    Set s2 = SortedSet.Create("a", "b", "c")
'
'    Dim nc As NameCall
'    Set nc = NameCall.OnObject("Contains", VbMethod, 2)
'
'    BatteryApplicable.Battery nc, s1, True
'    BatteryApplicable.Battery nc, s2, False
'
'End Sub

