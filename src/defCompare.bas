Attribute VB_Name = "defCompare"
'
' Compare
' -------
'
Public Enum CompareResult
    
    lt = -1
    eq = 0
    gt = 1

End Enum
Private Sub CheckComparable(ByVal x)
    
    If Not IsComparable(x) Then
        Dim msg As String
        msg = defShow.Show(x) & " Is not a Comparable object"
        TypeError "defCompare", "CheckComparable", msg
    End If
    
End Sub
Private Function IsComparable(ByVal x) As Boolean

    IsComparable = (TypeOf x Is Comparable)

End Function
'
' Public Functions
' ----------------
'
Public Function AsComparable(ByVal x) As Comparable
    
    CheckComparable x
    Set AsComparable = x
    
End Function
Public Function Compare(ByVal x, ByVal y) As CompareResult

    
    Dim result As CompareResult
    If Equals(x, y) Then
        result = eq
    ElseIf LessThan(x, y) Then
        result = lt
    ElseIf GreaterThan(x, y) Then
        result = gt
    End If
    
    Compare = result
    
End Function
Public Function LessThan(ByVal x, ByVal y) As Boolean

    Dim result As Boolean

    If IsComparable(x) Then
        result = (AsComparable(x).Compare(y) = lt)
    ElseIf IsComparable(y) Then
        result = (AsComparable(y).Compare(x) = gt)
    Else
        On Error GoTo ErrHandler
        result = (x < y)
        On Error GoTo 0
    End If
    
    LessThan = result
Exit Function
ErrHandler:
    Select Case Err.Number
    Case Else
        Exceptions.BubbleError "defCompare", "LessThan", Err
    End Select
End Function
Public Function GreaterThan(ByVal x, ByVal y) As Boolean

    GreaterThan = LessThan(y, x)
    
End Function
Public Function LessThanOrEqualTo(ByVal x, ByVal y) As Boolean

    LessThanOrEqualTo = Not GreaterThan(x, y)
    
End Function
Public Function GreaterThanOrEqualTo(ByVal x, ByVal y) As Boolean

    GreaterThanOrEqualTo = Not LessThan(x, y)
    
End Function




