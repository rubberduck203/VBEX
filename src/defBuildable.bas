Attribute VB_Name = "defBuildable"
Option Explicit

Public Function Repeat(ByVal seed As Buildable, ByVal val, ByVal n As Long) _
        As Buildable

    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim i As Long
    For i = 1 To n
        result.AddItem val
    Next

    Set Repeat = result

End Function
Public Function Enumerate(ByVal seed As Buildable, ByVal from As Long, _
        ByVal til As Long, Optional ByVal by As Long = 1) As Buildable

    If Not (0 < (til - from) * Sgn(by)) Then ' Does not converge
        Exceptions.ValueError seed, "Enumerate", "Sequence does not converge"
    End If

    Dim result As Buildable
    Set result = seed.MakeEmpty

    Dim i As Long
    For i = from To til Step by
        result.AddItem i
    Next

    Set Enumerate = result

End Function
''
' Converts an Transversable to any Buildable
Public Function ConvertTo(ByVal seed As Buildable, ByVal transversable) _
        As Variant
        
    Dim result As Buildable
    Set result = seed.MakeEmpty
    result.AddItems transversable
    Set ConvertTo = result

End Function
