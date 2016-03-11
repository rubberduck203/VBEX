Attribute VB_Name = "defSetLike"
Option Explicit
'
' Comparison
' ----------
'
Public Function SetEquals(ByVal xs As SetLike, ByVal ys) As Boolean

    If TypeOf ys Is SetLike Then
        On Error GoTo Nope
        SetEquals = (xs.Difference(ys).Count = 0)
        On Error GoTo 0
    Else
    
CleanExit:
        SetEquals = False
    End If
Exit Function
Nope:
    Resume CleanExit:
    
End Function
Public Function IsDisJoint(ByVal xs As SetLike, ByVal ys As SetLike) As Boolean

    IsDisJoint = (xs.Intersect(ys).Count = 0)
    
End Function
Public Function IsSubSetOf(ByVal xs As SetLike, ByVal ys As SetLike) As Boolean
    
    Dim x
    For Each x In xs
    
        If Not ys.Contains(x) Then
        
            IsSubSetOf = False
            Exit Function
            
        End If
        
    Next
    
    IsSubSetOf = True
    
End Function
Public Function IsProperSubSetOf(ByVal xs As SetLike, ByVal ys As SetLike) As Boolean

    IsProperSubSetOf = (xs.IsSubSetOf(ys) And (xs.Count < ys.Count))
    
End Function
Public Function IsSuperSetOf(ByVal xs As SetLike, ByVal ys As SetLike) As Boolean

    IsSuperSetOf = ys.IsSubSetOf(xs)
    
End Function
Public Function IsProperSuperSetOf(ByVal xs As SetLike, ByVal ys As SetLike) As Boolean

    IsProperSuperSetOf = ys.IsProperSubSetOf(xs)
    
End Function
'
' Constructors
' ------------
'
Public Function Union(ByVal seed As Buildable, ByVal xs, ByVal ys) As Variant

    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    result.AddItems xs
    result.AddItems ys
    
    Set Union = result
    
End Function
Public Function Intersect(ByVal seed As Buildable, ByVal xs, _
        ByVal ys As SetLike) As Variant

    Set Intersect = GenericJoin(True, seed, xs, ys)
    
End Function
Public Function Difference(ByVal seed As Buildable, ByVal xs, _
        ByVal ys As SetLike) As Variant
    
    Set Difference = GenericJoin(False, seed, xs, ys)
    
End Function
Public Function SymmetricDifference(ByVal seed As Buildable, _
        ByVal xs As SetLike, ByVal ys As SetLike) As Variant

    Dim leftOuter
    Set leftOuter = xs.Difference(ys)
    
    Dim rightOuter
    Set rightOuter = ys.Difference(xs)
    
    Set SymmetricDifference = Union(seed, leftOuter, rightOuter)
    
End Function
Private Function GenericJoin(ByVal contained As Boolean, _
        ByVal seed As Buildable, ByVal xs, ByVal ys As SetLike) As Variant

    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim x
    For Each x In xs
        If ys.Contains(x) = contained Then
            result.AddItem x
        End If
    Next
    
    Set GenericJoin = result

End Function
