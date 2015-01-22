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
' IPrintable
' ----------
'
Public Function ToString(ByVal x As Variant) As String

    If IsObject(x) Then
        If TypeOf x Is IPrintable Then
            ToString = x.ToString
            Exit Function
        End If
    End If
    
On Error GoTo NoDefaultProperty
    ToString = CStr(x)
    
Exit Function

NoDefaultProperty:

    RaiseNoDefaultProperty Err, "ToString", x
    ReRaiseError Err, "ToString"
    
End Function
'
'
''
' Not actually for `IPrintable` but sequence objects to
' use.
Public Function SequenceToString(ByVal xs As ISequence, _
        Optional ByVal delim As String, _
        Optional ByVal lcap As String = "(", _
        Optional ByVal rcap As String = ")") As String

    Dim ss() As Variant
    ss = xs.ToArray
    
    Dim i As Long
    For i = LBound(ss) To UBound(ss)
        ss(i) = ToString(ss(i))
    Next i
    
    SequenceToString = TypeName(xs) & lcap & Join(ss, delim) & rcap
    
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
    
    Dim xIsObj As Boolean
    xIsObj = IsObject(x)
    
    Dim yIsObj As Boolean
    yIsObj = IsObject(y)
    
    Dim xIsEquatable As Boolean
    If xIsObj Then
        xIsEquatable = TypeOf x Is IEquatable
    Else
        xIsEquatable = False
    End If
    
    Dim yIsEquatable As Boolean
    If yIsObj Then
        yIsEquatable = TypeOf y Is IEquatable
    Else
        yIsEquatable = False
    End If
    
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
