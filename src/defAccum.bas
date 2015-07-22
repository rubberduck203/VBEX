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
Public Function Scan(ByVal seed As IBuildable, ByVal op As IApplicable, ByVal init, ByVal sequence) As IBuildable

    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim temp
    Assign temp, init
    
    Dim element
    For Each element in sequence
    
        Assign temp, op.Apply(temp, element)
        result.AddItem temp
        
    Next
    
    Set Scan = result

End Function
