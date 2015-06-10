Attribute VB_Name = "IterableErrors"
Option Explicit
'
' Error to assist IIterable Classes
'
Public Sub CheckIndexRange(ByVal iterable As IIterable, ByVal method As String, ByVal index As Long)

    If index < iterable.LowerBound Then

        LowerThanLowerBoundError iterable, method, index

    ElseIf index > iterable.UpperBound Then
    
        GreaterThanUpperBoundError iterable, method, index
        
    End If
    
End Sub
Private Sub LowerThanLowerBoundError(ByVal iterable As IIterable, _
        ByVal method As String, ByVal index As Long)

    Dim msg As String
    msg = "Index " & index & " is lower than lowerbound: " & iterable.LowerBound
    IndexError iterable, method, msg

End Sub
Private Sub GreaterThanUpperBoundError(ByVal iterable As IIterable, _
        ByVal method As String, ByVal index As Long)

    Dim msg As String
    msg = "Index " & index & " is greater than upperbound: " & iterable.UpperBound
    IndexError iterable, method, msg

End Sub
