Attribute VB_Name = "defMap"
Option Explicit

Public Function Map(ByVal seed As IBuildable, _
        ByVal op As IApplicable, ByVal sequence) As Variant
    
    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        result.AddItem op.Apply(element)
    Next
    
    Set Map = result
    
End Function
Public Function FlatMap(ByVal seed As IBuildable, _
        ByVal op As IApplicable, ByVal sequence) As Variant
    
    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        result.AddItems op.Apply(element)
    Next
    
    Set FlatMap = result
    
End Function

