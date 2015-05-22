Attribute VB_Name = "defFold"
Option Explicit

Public Function Fold(ByVal op As IApplicable, ByVal init, ByVal sequence)

    Dim result
    Assign result, init
    
    Dim element
    For Each element In sequence
        Assign result, op.Apply(result, element)
    Next
    
    Assign Fold, result
    
End Function
