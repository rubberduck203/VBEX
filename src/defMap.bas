Attribute VB_Name = "defMap"
Option Explicit
'
' defMap
' ======
'
' Default implementations of Map and Bind for different structures. Classes 
' that implement `Map`, or `Bind` (aka `FlatMap`), can avoid code duplication
' by using these functions with their predeclared object as `seed`.
'
' TODO: Should these belong in IBuildable?
Private Const MAP_ADD As String = "AddItem"
Private Const BIND_ADD As String = "AddItems"
' Transversable
' -------------
' 
' Transversable maps use `For Each` structures to loop over the sequence. Since
' mutliple data-types can use `For Each` the squence is a Variant.
'
' These will be the most commonly used.
'
Public Function TransversableMap(ByVal seed As IBuildable, _
        ByVal op As IApplicable, ByVal sequence) As IBuildable
    
    On Error Goto Bubble
    Set TransversableMap = GenericTransversableMap(MAP_ADD, seed, op, sequence)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "TransversableMap", Err
    
End Function
Public Function TransversableBind(ByVal seed As IBuildable, _
        ByVal op As IApplicable, ByVal sequence) As IBuildable
    
    On Error Goto Bubble
    Set TransversableBind = GenericTransversableMap(BIND_ADD, seed, op, sequence)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "TransversableBind", Err
    
End Function
Private Function GenericTransversableMap(ByVal buildMethod As String, _
        ByVal seed As IBuildable, ByVal op As IApplicable, ByVal sequence) As IBuildable
    
    Dim result As IBuildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        CallByName result, buildMethod, vbMethod, op.Apply(element)
    Next
    
    Set GenericTransversableMap = result
    
End Function
'
' Iterable
' --------
'
' Use for any iterable classes that are not transversable.
' Result must still be buildable.
'
Public Function IterableMap(ByVal seed As IBuildable, ByVal op As IApplicable, _
        ByVal iterable As IIterable) AS IBuildable
    
    On Error Goto Bubble
    Set IterableMap = GenericIterableMap(MAP_ADD, seed, op, iterable)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "IterableMap", Err
    
End Function
Public Function IterableBind(ByVal seed As IBuildable, ByVal op As IApplicable, _
        ByVal iterable As IIterable) AS IBuildable
    
    On Error Goto Bubble
    Set IterableMap = GenericIterableMap(BIND_ADD, seed, op, iterable)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "IterableBind", Err
    
End Function
Private Function GenericIterableMap(ByVal buildMethod As String, _
        ByVal seed As IBuildable, ByVal op As IApplicable, _
        ByVal iterable As IIterable) AS IBuildable
    
    Dim result As IBuildable
    Set result = IBuildable.MakeEmpty
    
    Dim i As Long
    For i = iterable.LowerBound To iterable.UpperBound
         CallByName result, buildMethod, vbMethod, op.Apply(iterable.Item(i))
    Next
    
    Set GenericIterableMap = result
    
End Function








