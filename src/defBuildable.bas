Attribute VB_Name = "defBuildable"
Option Explicit

Public Function Repeat(ByVal seed As IBuildable, ByVal val, ByVal n As Long) As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim i As Long
    For i = 1 To n
        result.AddItem val
    Next

    Set Repeat = result

End Function
