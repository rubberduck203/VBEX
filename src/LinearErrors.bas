Attribute VB_Name = "LinearErrors"
Option Explicit
'
' Error to assist Linear Classes
'
Public Sub CheckIndexRange(ByVal sequence As Linear, ByVal method As String, ByVal index As Long)

    If index < sequence.LowerBound Then

        LowerThanLowerBoundError sequence, method, index

    ElseIf index > sequence.UpperBound Then
    
        GreaterThanUpperBoundError sequence, method, index
        
    End If
    
End Sub
Private Sub LowerThanLowerBoundError(ByVal sequence As Linear, _
        ByVal method As String, ByVal index As Long)

    Dim msg As String
    msg = "Index " & index & " is lower than lowerbound: " & sequence.LowerBound
    IndexError sequence, method, msg

End Sub
Private Sub GreaterThanUpperBoundError(ByVal sequence As Linear, _
        ByVal method As String, ByVal index As Long)

    Dim msg As String
    msg = "Index " & index & " is greater than upperbound: " & sequence.UpperBound
    IndexError sequence, method, msg

End Sub
