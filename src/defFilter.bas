Attribute VB_Name = "defFilter"
Option Explicit

Public Function Filter(ByVal seed As Buildable, _
        ByVal pred As Applicable, ByVal sequence) As Variant
    
    Set Filter = GenericFilter(True, seed, pred, sequence)
    
End Function
Public Function FilterNot(ByVal seed As Buildable, _
        ByVal pred As Applicable, ByVal sequence) As Variant
    
    Set FilterNot = GenericFilter(False, seed, pred, sequence)
    
End Function
Private Function GenericFilter(ByVal keep As Boolean, _
        ByVal seed As Buildable, ByVal pred As Applicable, _
        ByVal sequence) As Variant
        
    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        If pred.Apply(element) = keep Then
            result.AddItem element
        End If
    Next
    
    Set GenericFilter = result
        
End Function


