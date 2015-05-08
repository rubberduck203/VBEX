Attribute VB_Name = "defFilter"
Option Explicit

Public Function Filter(ByVal seed As IBuildable, _
        ByVal pred As IApplicable, ByVal sequence) As Variant
    
    Set Filter = GenericFilter(True, seed, pred, sequence)
    
End Function
Public Function FilterNot(ByVal seed As IBuildable, _
        ByVal pred As IApplicable, ByVal sequence) As Variant
    
    Set FilterNot = GenericFilter(False, seed, pred, sequence)
    
End Function
Private Function GenericFilter(ByVal keep As Boolean, _
        ByVal seed As IBuildable, ByVal pred As IApplicable, _
        ByVal sequence) As Variant
        
    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        If pred.Apply(element) = keep Then
            result.AddItem element
        End If
    Next
    
    Set GenericFilter = result
        
End Function


