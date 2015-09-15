Attribute VB_Name = "cast"
Option Explicit
'
' cast
' ====
'
''
' Assign x to y regardless of object or primitive
Public Sub Assign(ByRef x, ByVal y)

    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If

End Sub
''
' Convert Variant To Varaint()
' TODO: Bubble Errors if xs is not an array
Public Function CArray(ByVal xs) As Variant()

On Error GoTo CheckIfNotArray
    CArray = xs
    
Exit Function
CheckIfNotArray:
    Debug.Assert Not IsArray(xs)
    NotImplementedError "cast", "CArray"
    
End Function
''
' Is x An Array?
Public Function IsArray(ByVal x) As Boolean

    IsArray = (TypeName(x) Like "*()")

End Function

