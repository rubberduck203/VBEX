VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' List Class
' ==========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (List.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Implements Equatable
Implements Linear
Implements Buildable
Implements Showable
Implements Monadic
'
' Private Members
' ---------------
'
Private pCollec As Collection ' Sole datamember
'
' Public Members
' --------------
'
' None
'
' Private Methods
' ---------------
'
Private Sub TransformIndex(ByRef x As Variant)

    If x < 0 Then x = x + pCollec.Count + 1
    
End Sub
''
' Replace
Private Sub Replace(ByVal index As Long, ByVal element As Variant)
    
    pCollec.Remove index
    
    If index = pCollec.Count + 1 Then
        pCollec.Add element
    Else
        pCollec.Add element, Before:=index
    End If
    
End Sub
Private Sub RaiseNotSequenceError(ByVal method As String, ByVal scalar As Variant)

    Dim msg As String
    msg = """" & TypeName(scalar) & """ is not an iterable type."
    
    TypeError Me, method, msg
    
End Sub
'
''
' Constructors
' ------------
'
Public Function Copy(ByVal sequence As Variant) As List
Attribute Copy.VB_Description = "Create a new List using the elements of `sequence`."

    Dim result As New List
    result.Extend sequence
    Set Copy = result
    
End Function
Public Function Create(ParamArray args() As Variant) As List
Attribute Create.VB_Description = "Create a new List using the elements of `args`."

    Set Create = Copy(CVar(args))
    
End Function
Public Function Repeat(ByVal x, ByVal n As Long) As List
Attribute Repeat.VB_Description = "Create a new List of `x` repeated `n` times."

    Set Repeat = defBuildable.Repeat(List, x, n)

End Function
Public Function Enumerate(ByVal from As Long, ByVal til As Long, _
    Optional ByVal by As Long = 1) As List

    Set Enumerate = defBuildable.Enumerate(List, from, til, by)

End Function
'
' Class
' -----
'
Private Sub Class_Initialize()
    Set pCollec = New Collection
End Sub
''
' Destructor: Might be unnecessary
Private Sub Class_Terminate()
    Set pCollec = Nothing
End Sub
''
' Enumerator: How we all would love to create our own...
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = pCollec.[_NewEnum]
    
End Property
'
' Public Methods
' --------------
'
' ### Immutable operations
'
' #### Combiners
'
Public Function Concat(ByVal other As List) As List

    Dim result As List
    Set result = List.Copy(Me)
    result.Extend other
    Set Concat = result

End Function
Public Function Appended(ParamArray elems() As Variant) As List

    Dim result As List
    Set result = List.Copy(Me)
    result.Extend CArray(elems)
    Set Appended = result

End Function
Public Function Prepended(ParamArray elems() As Variant) As List

    Dim result As List
    Set result = List.Copy(Me)
    result.Insert 1, CArray(elems)
    Set Prepended = result

End Function
'
' #### Monadic
'
Public Function Map(ByVal op As Applicable) As List
Attribute Map.VB_Description = "Return a new List by applying `op` to the elements of this list.  Identical to LINQ Select"
    
    Set Map = defMap.TransversableMap(List, op, Me)
    
End Function
Public Function FlatMap(ByVal op As Applicable) As List
Attribute FlatMap.VB_Description = "Return a new List by applying `op` to the elements of this list and concatenating the results."

    Set FlatMap = defMap.TransversableBind(List, op, Me)

End Function
'
' ### Filter
'
Public Function Filter(ByVal pred As Applicable) As List
Attribute Filter.VB_Description = "Return a new List of only the elements that satisfy `pred`.  Identical to LINQ Where"
    
    Set Filter = defFilter.Filter(List, pred, Me)
    
End Function
Public Function FilterNot(ByVal pred As Applicable) As List
Attribute FilterNot.VB_Description = "Return a new List of only the elements that fail `pred`."
    
    Set FilterNot = defFilter.FilterNot(List, pred, Me)
    
End Function
'
' #### Accumulators
'
Public Function Fold(ByVal Init As Variant, ByVal accum As Applicable) As Variant
Attribute Fold.VB_Description = "Accumulates a value by applying `accum` to each value using a given intial value. `Tail.Fold(Head,f) = Reduce(f)`"
    
    Assign Fold, defAccum.Fold(accum, Init, Me)
    
End Function
Public Function Reduce(ByVal accum As Applicable) As Variant
Attribute Reduce.VB_Description = "Accumulates a value by applying `accum` to each value using the first element as the initial value. Identical to LINQ Aggregate"
    
    Assign Reduce, Tail.Fold(Head, accum)
    
End Function
Public Function Scan(ByVal Init As Variant, ByVal accum As Applicable) As List
Attribute Scan.VB_Description = "Accumulates a list of by folding with `accum` and storing each result. `Scan(z,f).Last = Fold(z,f)`"
    
    Set Scan = defAccum.Scan(List, accum, Init, Me)
    
End Function
'
' #### Sectioning
'
Public Function Head() As Variant
Attribute Head.VB_Description = "Returns the first element. Identical to LINQ First"
    
    Assign Head, Item(1)
    
End Function
Public Function MaybeHead() As Maybe
Attribute MaybeHead.VB_Description = "Returns the first element if the list is not empty."
    
    If Not IsEmpty Then
        Set MaybeHead = Maybe.Some(Head)
    Else
        Set MaybeHead = Maybe.None
    End If
    
End Function
Public Function Tail() As List
Attribute Tail.VB_Description = "Returns a new list of all but the first element."
    
    Dim result As List
    If Count > 1 Then
        Set result = Slice(2, Count)
    Else
        Set result = Create()
    End If
    
    Set Tail = result
    
End Function
Public Function Last() As Variant
Attribute Last.VB_Description = "Returns the last element."
    
    Assign Last, Item(Count)
    
End Function
Public Function MaybeLast() As Variant
Attribute MaybeLast.VB_Description = "Returns the last element if the list is not empty"
    
    If Not IsEmpty Then
        Set MaybeLast = Maybe.Some(Last)
    Else
        Set MaybeLast = Maybe.None
    End If
    
End Function
Public Function Init() As List

    Set Init = Slice(1, Count - 1)

End Function
'
' #### Misc.
'
Public Function Reverse() As List
Attribute Reverse.VB_Description = "Returns the reverse of the list."
    
    Set Reverse = Slice(Count, 1)
    
End Function
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of elements."
    
    Count = pCollec.Count
    
End Function
Public Function IsEmpty() As Boolean

    IsEmpty = (Linear_UpperBound < Linear_LowerBound)

End Function
''
' GroupBy returns a Dict until Dict has an interface
Public Function GroupBy(ByVal op As Applicable) As Dict
Attribute GroupBy.VB_Description = "Groups elements into a Dict according to `op`."

    Dim result As Dict
    Set result = Dict.Create

    Dim element
    For Each element In pCollec
        
        Dim key
        Assign key, op.Apply(element)
        
        If Not result.Exists(key) Then
            result.Add key, List.Create()
        End If
        
        result.Item(key).Append element
        
    Next

    Set GroupBy = result

End Function
' TODO: Additional methods below
Public Function Drop(ByVal n As Long) As List

    Dim result As List
    If n < Count Then
        Set result = Slice(n + 1, Count)
    Else
        Set result = Create()
    End If

    Set Drop = result

End Function
' TODO: Additional methods below
Public Function Take(ByVal n As Long) As List

    Dim result As List
    If n < Count Then
        Set result = Slice(1, n)
    Else
        Set result = Copy(Me)
    End If

    Set Take = result

End Function
' FoldLeft, FoldRight, ReduceLeft, ReduceRight
Public Function IndexWhere(ByVal pred As Applicable) As Maybe

    Dim result As Maybe
    Set result = Maybe.None
    
    Dim i As Long
    i = Linear_LowerBound
    
    Do While result.IsNone And i <= Linear_UpperBound
        Set result = Maybe.MakeIf(pred.Apply(Item(i)), i)
        i = i + 1
    Loop
    
    Set IndexWhere = result

End Function
Public Function IndexOf(ByVal val) As Maybe

    Set IndexOf = IndexWhere(InternalDelegate.Make("Equals").Partial(val, Empty))
    
End Function
Public Function LastIndexWhere(ByVal pred As Applicable) As Maybe

    Dim indexFromBack As Maybe
    Set indexFromBack = Reverse.IndexWhere(pred)
    
    Dim fromFront As Applicable
    Set fromFront = Lambda.FromShort(Linear_UpperBound + 1 & " - _")
    
    Set LastIndexWhere = indexFromBack.Map(fromFront)
    
End Function
Public Function LastIndexOf(ByVal val) As Maybe

    Dim equalsVal As Applicable
    Set equalsVal = InternalDelegate.Make("Equals").Partial(val, Empty)
    
    Set LastIndexOf = LastIndexWhere(equalsVal)
    
End Function
Public Function Find(ByVal pred As Applicable) As Maybe

    Set Find = IndexWhere(pred).Map(OnArgs.Make("Item", VbGet, Me))
    
End Function
Public Function FindLast(ByVal pred As Applicable) As Maybe

    Set FindLast = LastIndexWhere(pred).Map(OnArgs.Make("Item", VbGet, Me))

End Function
Public Function CountWhere(ByVal pred As Applicable) As Long

    CountWhere = Filter(pred).Count

End Function
' Tails() As Iterator[List] whenever Iterator exists
' Inits() As Iterator[List]
' Sliding(???) As Iterator[List]
'
' ### Mutators
'
' General pattern is to have a property or method that applies to one
' element and then extend that to sequences.
'
' #### Access and Replace
'
' `Item` allows for access and replacement of an element through indexes.
' `Slice` extends that to sequences that allows reversing and steps. Only
' `Slice has steps, other sequence methods only apply to continuous sequences.
'
''
' Default member is Item
Public Property Let Item(ByVal index As Long, ByVal val As Variant)
Attribute Item.VB_Description = "Replaces the element at `index` with `val`."
Attribute Item.VB_UserMemId = 0
    
    TransformIndex index
    Replace index, val
    
End Property
Public Property Set Item(ByVal index As Long, ByVal val As Object)
Attribute Item.VB_Description = "Replaces the element at `index` with `val`."
Attribute Item.VB_UserMemId = 0
    
    TransformIndex index
    Replace index, val
    
End Property
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_Description = "Returns the element at `index`."
Attribute Item.VB_UserMemId = 0
    
    TransformIndex index
    cast.Assign Item, pCollec.Item(index)
    
End Property
Public Property Get Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1) As List
    
    TransformIndex a
    TransformIndex b
    
    Set Slice = New List
    
    If s < 1 Then
        ValueError Me, "Slice", "Step " & s & " is not a natural number."
    End If
    
    s = IIf(a < b, s, -s)
    
    Dim i As Long
    For i = a To b Step s
        Slice.Append pCollec.Item(i)
    Next i
    
End Property
Public Property Let Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1, ByVal sequence As Variant)
    
    TransformIndex a
    TransformIndex b
    
    If s < 1 Then
        ValueError Me, "Slice", "Step " & s & " is not a natural number."
    End If
    s = IIf(a < b, s, -s)
    
    Dim other As List
    Set other = List.Copy(sequence)
    
    If Abs(a - b) + 1 <> other.Count Then
        Err.Raise 9, TypeName(Me) & ".Slice", "Subscript out of Range."
    End If
    
    Dim i As Long: i = a
    
    On Error GoTo NotSequence
    Dim element As Variant
    For Each element In other
        
        Replace i, element
        i = i + s
        
    Next element
    
    Debug.Assert (i - s = b)
    
Exit Property
NotSequence:
    RaiseNotSequenceError "Slice", sequence
    
End Property
'
' #### Removal
'
''
' Remove
Public Sub Remove(ByVal index As Long)
Attribute Remove.VB_Description = "Removes the element at `index`."
    
    TransformIndex index
    pCollec.Remove index
    
End Sub
''
' Revisions from CR
Public Sub Clear(ByVal start As Long, ByVal size As Long)
Attribute Clear.VB_Description = "Removes `size` elements at `start`."
    
    TransformIndex start
    
    Dim i As Long
    For i = 1 To size
        pCollec.Remove start
    Next i

End Sub
'
' #### Extension
'
''
' Append: Vanilla `Add`
Public Sub Append(ParamArray elems() As Variant)
Attribute Append.VB_Description = "Adds `elems` to the end of the list."

    Extend CVar(elems)
    
End Sub
Public Sub Prepend(ParamArray elems() As Variant)

    Insert 1, elems
    
End Sub
''
' Extend: Append sequence
Public Sub Extend(ByVal sequence As Variant)
Attribute Extend.VB_Description = "Adds the elements of `sequence` to the end of the list."
    
    On Error GoTo NotSequence
    Dim element As Variant
    For Each element In sequence
        pCollec.Add element
    Next element

Exit Sub
NotSequence:
    RaiseNotSequenceError "Extend", sequence
    
End Sub
'
' #### Insert
'
''
' Emplace: Place element at index, existing elements shifted backward.
Public Sub Emplace(ByVal index As Long, ParamArray elems() As Variant)
Attribute Emplace.VB_Description = "Inserts `elems` starting at `index`, shifting existing elements backward."
    
    Insert index, CVar(elems)
    
End Sub
''
' Insert: Sames as Emplace but with a sequence
Public Sub Insert(ByVal index As Long, ByVal sequence As Variant)
Attribute Insert.VB_Description = "Inserts the elements of `sequence` starting at `index`, shifting existing elements backward."

    TransformIndex index
    On Error GoTo NotSequence
    
    Dim other As List
    Set other = List.Copy(sequence).Reverse
    
    Dim element As Variant
    For Each element In other
        ' push back previous element
        pCollec.Add element, Before:=index
    Next element

Exit Sub
NotSequence:
    RaiseNotSequenceError "Insert", sequence
    
End Sub
'
' Interfaces
' ----------
'
' ### Buildable
'
Private Function Buildable_MakeEmpty() As Buildable

    Set Buildable_MakeEmpty = List.Create()

End Function
Private Sub Buildable_AddItem(ByVal element)

    Append element

End Sub
Private Sub Buildable_AddItems(ByVal sequence)

    Extend sequence

End Sub
'
' ### Linear
'
Private Function Linear_Item(ByVal index As Long) As Variant

    Assign Linear_Item, Item(index)

End Function
Private Function Linear_LowerBound() As Long

    Linear_LowerBound = 1

End Function
Private Function Linear_UpperBound() As Long

    Linear_UpperBound = Count

End Function
Public Function ToArray() As Variant()

    ToArray = defIterable.ToArray(Me)

End Function
Private Function Linear_ToArray() As Variant()

    Linear_ToArray = ToArray

End Function
Public Function ToCollection() As Collection

    Set ToCollection = defIterable.ToCollection(Me)

End Function
Private Function Linear_ToCollection() As Collection

    Set Linear_ToCollection = ToCollection
    
End Function
Public Function ToBuildable(ByVal seed As Buildable) As Buildable

    Set ToBuildable = defIterable.ToBuildable(seed, Me)

End Function
Private Function Linear_ToBuildable(ByVal seed As Buildable) As Buildable

    Set Linear_ToBuildable = ToBuildable(seed)

End Function
'
' ### Equatable
'
Public Function Equals(ByVal x) As Boolean
Attribute Equals.VB_Description = "Check this list is identical to unknown value `x`."

    Equals = False
    
    If TypeName(Me) <> TypeName(x) Then Exit Function
    
    Dim other As List
    Set other = x
    
    If other.Count <> Count Then Exit Function
    
    Dim i As Long
    For i = 1 To Count
    
        Dim a, b
        Assign a, Item(i)
        Assign b, other.Item(i)
        
        If Not defEquals.Equals(a, b) Then Exit Function
    Next
    
    Equals = True
    
End Function
Private Function Equatable_Equals(ByVal other) As Boolean

    Equatable_Equals = Equals(other)
    
End Function
'
' ### Monadic
'
Private Function Monadic_Bind(ByVal op As Applicable) As Monadic
    
    Set Monadic_Bind = FlatMap(op)
    
End Function
Private Function Monadic_Map(ByVal op As Applicable) As Monadic
    
    Set Monadic_Map = Map(op)
    
End Function
Private Function Monadic_Unit(ByVal x) As Monadic
    
    Set Monadic_Unit = List.Create(x)
    
End Function
'
'  ### Showable
'
Public Function Show() As String
Attribute Show.VB_Description = "Represent the List as a String."
    
    Show = defShow.ShowableObject(Me, ToArray)
    
End Function
Private Function Showable_Show() As String
    
    Showable_Show = Show()
    
End Function
