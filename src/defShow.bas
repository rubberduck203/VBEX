Attribute VB_Name = "defShow"
Option Explicit
'
' Show
' ====
'
' Default implementation of Show.
'
Private Const ARRAY_DELIM As String = ", "
Private Const SQUARE_ARRAY_DELIM As String = "; "
Private Const OBJOPEN As String = "("
Private Const OBJCLOSE As String = ")"
Private Const ARROPEN As String = "["
Private Const ARRCLOSE As String = "]"

'
' Public Methods
' --------------
'
Public Function Show(ByVal x) As String
    Dim result As String

    If TypeOf x Is Showable Then
        
        Dim s As Showable
        Set s = x
        result = s.Show
        
    ElseIf IsObject(x) Then
        result = UnShowableObject(x)
    ElseIf cast.IsArray(x) Then
        result = ShowArray(x)
    ElseIf IsNull(x) Then
        x = vbNullString
    Else
        result = x
    End If
    
    Show = result

End Function
Public Function ParamShowableObject(ByVal obj As Object, _
        ParamArray members()) As String
        
    ParamShowableObject = ShowableObject(obj, cast.CArray(members))
        
End Function
Public Function ShowableObject(ByVal obj As Object, ByRef members()) As String

    Dim shownMembers As String
    shownMembers = ShowArrayMembers(members)
    
    ShowableObject = TypeName(obj) & OBJOPEN & shownMembers & OBJCLOSE

End Function
'
' Private Methods
' ---------------
'
Private Function UnShowableObject(ByVal obj As Object) As String

    Dim repr As String
    repr = "&" & ObjPtr(obj)

    UnShowableObject = ShowableObject(obj, cast.xArray(repr))

End Function
'
' ### Showing Arrays
'
Private Function ShowArray(ByRef xs As Variant) As String

    Dim shownMembers As String
    If IsSquareArray(xs) Then
        shownMembers = ShowSquareArrayMembers(xs)
    Else
        shownMembers = ShowArrayMembers(xs)
    End If
    
    Dim withParens As String
    withParens = TypeName(xs)
    
    Dim withoutParens As String
    withoutParens = Left(withParens, Len(withParens) - 2)
    
    ShowArray = withoutParens & ARROPEN & shownMembers & ARRCLOSE
    
End Function
Private Function IsSquareArray(ByRef xs As Variant) As Boolean

    Dim dummy As Long
    Dim result As Boolean
    
    On Error GoTo Nope
    dummy = UBound(xs, 2)
    
    On Error GoTo Yup
    dummy = UBound(xs, 3)
    
    On Error GoTo 0
    Exceptions.TypeError "defShow", "IsSquareArray", _
        "Can not and will not show 3 or more dimensional array." & _
        "  Do not use cubic or greater arrays!"

CleanExit:
    IsSquareArray = result
Exit Function

Nope:
    Err.Clear
    result = False
    Resume CleanExit
    
Yup:
    Err.Clear
    result = True
    Resume CleanExit
    
End Function
Private Function ShowArrayMembers(ByRef xs As Variant) As String

    Dim lower As Long
    lower = LBound(xs)
    
    Dim upper As Long
    upper = UBound(xs)
    
    Dim results() As String
    If lower <= upper Then
        ReDim results(lower To upper)
    End If
    
    Dim i As Long
    For i = lower To upper
        results(i) = Show(xs(i))
    Next i

    ShowArrayMembers = Join(results, ARRAY_DELIM)

End Function
Private Function ShowSquareArrayMembers(ByRef xs As Variant) As String

    Dim txs As Variant
    txs = Application.Transpose(xs)

    Dim lower As Long
    lower = LBound(txs)
    
    Dim upper As Long
    upper = UBound(txs)
    
    Dim Size As Long
    Size = upper - lower + 1
    
    Dim results() As String
    If lower <= upper Then
        ReDim results(1 To Size)
    End If
    
    Dim i As Long
    For i = 1 To Size
        results(i) = ShowArrayMembers(Application.index(txs, i, 0))
    Next
    
    ShowSquareArrayMembers = Join(results, SQUARE_ARRAY_DELIM)
    
End Function

