Attribute VB_Name = "defAccum"
Option Explicit

Public Function Fold(ByVal op As Applicable, ByVal Init, ByVal sequence)
    
    Dim result
    Assign result, Init
    
    Dim element
    For Each element In sequence
        Assign result, op.Apply(result, element)
    Next
    
    Assign Fold, result
    
End Function
Public Function Scan(ByVal seed As Buildable, ByVal op As Applicable, ByVal Init, ByVal sequence) As Buildable

    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim temp
    Assign temp, Init
    
    Dim element
    For Each element In sequence
    
        Assign temp, op.Apply(temp, element)
        result.AddItem temp
        
    Next
    
    Set Scan = result

End Function
