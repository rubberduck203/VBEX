Attribute VB_Name = "defShow"
Option Explicit
'
' Show
' ====
'
' Default implementation of Show.
'
Private Const delim As String = ", "
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

    If TypeOf x Is IShowable Then
        
        Dim s As IShowable
        Set s = x
        result = s.Show
        
    ElseIf IsObject(x) Then
        result = UnShowableObject(x)
    ElseIf cast.IsArray(x) Then
        result = ShowArray(x)
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

    Dim shownMembers() As String
    shownMembers = ShowArrayMembers(members)
    
    ShowableObject = TypeName(obj) & OBJOPEN & Join(shownMembers, delim) & OBJCLOSE

End Function
'
' Private Methods
' ---------------
'
Private Function UnShowableObject(ByVal obj As Object) As String

    Dim repr As String
    repr = "&" & ObjPtr(obj)

    UnShowableObject = ShowableObject(obj, cast.CArray(Array(repr)))

End Function
Private Function ShowArray(ByRef xs) As String

    Debug.Assert IsArray(xs)
    
    Dim shownMembers() As String
    shownMembers = ShowArrayMembers(xs)
    
    Dim withParens As String
    withParens = TypeName(xs)
    
    Dim withoutParens As String
    withoutParens = Left(withParens, Len(withParens) - 2)
    
    ShowArray = withoutParens & ARROPEN & Join(shownMembers, delim) & ARRCLOSE
    
End Function
''
' Map `Show` on array xs
Private Function ShowArrayMembers(ByRef xs) As String()

    Debug.Assert IsArray(xs)

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

    ShowArrayMembers = results

End Function



