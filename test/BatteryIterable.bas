Attribute VB_Name = "BatteryIterable"
'@TestModule
Option Explicit
Option Private Module
Private Assert As New Rubberduck.AssertClass
'
' Iterable Battery
' ================
'
Public Sub Battery(ByVal itbl As Linear)

    LowerLTEQUpper itbl
    ItemInRange itbl
    ItemLTLower itbl
    ItemGTUpper itbl

End Sub
'
' Private Procedures
' ------------------
'
' ### Tests
'
Private Sub LowerLTEQUpper(ByVal itbl As Linear)

    Dim lower As Long
    lower = itbl.LowerBound

    Dim upper As Long
    upper = itbl.UpperBound

    Dim TestPass As Boolean
    TestPass = (lower <= upper)
    
    Dim msg As String
    msg = "Lower(" & lower & ") <= Upper(" & upper & ")" & _
    " For iterable(" & defshow.Show(itbl) & ")"
    
    Assert.IsTrue TestPass, msg

End Sub
Private Sub ItemInRange(ByVal itbl As Linear)

    Dim lower As Long
    lower = itbl.LowerBound

    Dim upper As Long
    upper = itbl.UpperBound
 
    Dim msg As String
    msg = "ItemInRange"

    Dim x
    On Error GoTo Fail
    Assign x, GetRandomItem(itbl, lower, upper)
    On Error GoTo 0
    
    Assert.IsFalse IsEmpty(x), msg

CleanExit:
Exit Sub
Fail:
    Assert.Fail msg
    Resume CleanExit
    
End Sub
Private Sub ItemLTLower(ByVal itbl As Linear)
    
    Dim lower As Long
    lower = itbl.LowerBound

    Dim msg As String
    msg = "ItemLTLower"

    Dim x
    On Error GoTo Pass
    Assign x, itbl.Item(lower - 1)
    On Error GoTo 0
    
    Assert.Fail msg
    
CleanExit:
Exit Sub
Pass:
    Assert.AreEqual Err.Number, CLng(9), msg
    
End Sub
Private Sub ItemGTUpper(ByVal itbl As Linear)

    Dim upper As Long
    upper = itbl.UpperBound
    
    Dim msg As String
    msg = "ItemGTUpper"

    Dim x
    On Error GoTo Pass
    Assign x, itbl.Item(upper + 1)
    On Error GoTo 0
    
    Assert.Fail msg
    
CleanExit:
Exit Sub
Pass:
    Assert.AreEqual Err.Number, CLng(9), msg
    
End Sub
'
' ### Helper Functions
'
Private Function GetRandomItem(ByVal itbl As Linear, ByVal lower As Long, _
        ByVal upper As Long) As Variant
    
    Dim ri As Long
    ri = RandomIndex(lower, upper)
    Assign GetRandomItem, itbl.Item(ri)
    
End Function
Private Function RandomIndex(ByVal lower As Long, ByVal upper As Long) As Long

    RandomIndex = Math.Round((upper - lower + 1) * Math.Rnd()) + lower

End Function
