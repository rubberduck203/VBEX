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
' It may be more prudent to split these into the various default files like
' defIterable, defTransversable etc...
'
' TODO: Should these belong in Buildable?
Private Const MAP_ADD As String = "AddItem"
Private Const BIND_ADD As String = "AddItems"
'
' Transversable
' -------------
'
' Transversable maps use `For Each` structures to loop over the sequence. Since
' mutliple data-types can use `For Each` the squence is a Variant.
'
' These will be the most commonly used.
'
Public Function TransversableMap(ByVal seed As Buildable, _
        ByVal op As Applicable, ByVal sequence) As Buildable
    
    On Error GoTo Bubble
    Set TransversableMap = GenericTransversableMap(MAP_ADD, seed, op, sequence)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "TransversableMap", Err
    
End Function
Public Function TransversableBind(ByVal seed As Buildable, _
        ByVal op As Applicable, ByVal sequence) As Buildable
    
    On Error GoTo Bubble
    Set TransversableBind = GenericTransversableMap(BIND_ADD, seed, op, sequence)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "TransversableBind", Err
    
End Function
Private Function GenericTransversableMap(ByVal buildMethod As String, _
        ByVal seed As Buildable, ByVal op As Applicable, ByVal sequence) As Buildable
    
    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim element
    For Each element In sequence
        CallByName result, buildMethod, VbMethod, op.Apply(element)
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
Public Function IterableMap(ByVal seed As Buildable, ByVal op As Applicable, _
        ByVal iterable As Linear) As Buildable
    
    On Error GoTo Bubble
    Set IterableMap = GenericIterableMap(MAP_ADD, seed, op, iterable)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "IterableMap", Err
    
End Function
Public Function IterableBind(ByVal seed As Buildable, ByVal op As Applicable, _
        ByVal iterable As Linear) As Buildable
    
    On Error GoTo Bubble
    Set IterableBind = GenericIterableMap(BIND_ADD, seed, op, iterable)
    
Exit Function
Bubble:
    Exceptions.BubbleError "defMap", "IterableBind", Err
    
End Function
Private Function GenericIterableMap(ByVal buildMethod As String, _
        ByVal seed As Buildable, ByVal op As Applicable, _
        ByVal iterable As Linear) As Buildable
    
    Dim result As Buildable
    Set result = seed.MakeEmpty
    
    Dim i As Long
    For i = iterable.LowerBound To iterable.UpperBound
         CallByName result, buildMethod, VbMethod, op.Apply(iterable.Item(i))
    Next
    
    Set GenericIterableMap = result
    
End Function

