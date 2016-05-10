Attribute VB_Name = "defEquals"
Option Explicit
Private Sub CheckEquatable(ByVal x)
    
    If Not IsEquatable(x) Then
        TypeError "defEquals", "CheckEquatable"
    End If
    
End Sub
Private Function IsEquatable(ByVal x) As Boolean

    IsEquatable = (TypeOf x Is Equatable)

End Function
'
' Public Functions
' ----------------
'
Public Function AsEquatable(ByVal x) As Equatable
    
    CheckEquatable x
    Set AsEquatable = x
    
End Function
Public Function Equals(ByVal x, ByVal y) As Boolean

    Dim result As Boolean

    If IsEquatable(x) Then
        result = AsEquatable(x).Equals(y)
    ElseIf IsEquatable(y) Then
        result = AsEquatable(y).Equals(x)
    Else
        On Error GoTo ErrHandler
        result = (x = y)
        On Error GoTo 0
    End If
    
    Equals = result
    
Exit Function
ErrHandler:
    Select Case Err.Number
    Case Else
        Exceptions.BubbleError "defEquals", "Equals", Err
    End Select
    
End Function
