Attribute VB_Name = "defBuildable"
Option Explicit

Public Function Repeat(ByVal seed As IBuildable, ByVal val, ByVal n As Long) 
        As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim i As Long
    For i = 1 To n
        result.AddItem val
    Next

    Set Repeat = result

End Function
Public Function Enumerate(ByVal seed As IBuildable, ByVal from As Long, _
        ByVal til As Long, Optional ByVal by As Long = 1) As IBuildable

    If Not (0 < (til - from) * Sgn(by)) Then ' Does not converge
        Exceptions.ValueError seed, "Enumerate", "Sequence does not converge"
    End If

    Dim result As IBuildable
    Set result = seed.MakeEmpty

    Dim i As Long
    For i = from To til Step by
        result.AddItem i
    Next

    Set Enumerate = result

End Function
''
' Converts an Transversable to any IBuildable
Public Function ConvertTo(ByVal seed As IBuildable, ByVal transversable) 
        As Variant
        
    Dim result As IBuildable
    Set result = IBuildable.MakeEmpty
    result.AddItems(transversable)
    Set ToBuildable = result

End Function
