Attribute VB_Name = "cast"
Option Explicit
'
' cast
' ====
'
' Type casting interfaces for Data Directed Programming
'
'
'
Public Sub Assign(ByRef x As Variant, ByVal y As Variant)
    
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
    
End Sub
'
'
'
'
Public Function CArray(ByVal xs As Variant) As Variant()

    CArray = xs
    
End Function
'
' ICloneable
' ----------
'
Public Function Clone(ByVal x As Variant) As Variant

    If IsObject(x) Then
        If TypeOf x Is ICloneable Then
            Set Clone = x.Clone
        Else
            Err.Raise 438, "cast.Clone", "Clone: " & TypeName(x) & " is not a cloneable object"
        End If
    Else
        Clone = x
    End If

End Function
'
' IEquateable
' -----------
'
Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Boolean

    Dim xIsEquatable As Boolean
    xIsEquatable = TypeOf x Is IEquatable
    
    Dim yIsEquatable As Boolean
    yIsEquatable = TypeOf y Is IEquatable
    
    If xIsEquatable And yIsEquatable Then
        Equals = x.Equals(y)
    ElseIf xIsEquatable Xor yIsEquatable Then
        Equals = False
    Else
        On Error GoTo NoDefaultProperty
        Equals = (x = y)
    End If

Exit Function

NoDefaultProperty:
' TODO: x or y is offensive?
    RaiseNoDefaultProperty Err, "Equals", x
    ReRaiseError Err, "Equals"
    
End Function
'
' Errors
' ------
'
Private Function GetErrorSoruce(ByVal method As String) As String
    
    GetErrorSoruce = "cast." & method
    
End Function
Private Sub ReRaiseError(ByVal e As ErrObject, ByVal method As String)

    Err.Raise e.Number, GetErrorSoruce(method), e.Description, e.HelpFile, e.HelpContext
    
End Sub
Private Sub RaiseNoDefaultProperty(ByVal e As ErrObject, _
        ByVal method As String, ByVal obj As Variant)

    If e.Number = 438 Then
    
        Err.Raise 438, GetErrorSoruce(method), _
            "Class " & TypeName(obj) & " does not have a defualt property."
            
    ElseIf e.Number = 450 Then
    
        Err.Raise 438, GetErrorSoruce(method), _
            "Default property of " & TypeName(obj) & " is not nullary."
            
    End If
    
End Sub
