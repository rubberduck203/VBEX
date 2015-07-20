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
' This file (List.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
Implements IEquatable
Implements IIterable
Implements IBuildable
Implements IShowable
Implements IMonadic
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
' ### Constructors
'
Public Function Copy(ByVal sequence As Variant) As List

    Dim result As New List
    result.Extend sequence
    Set Copy = result
    
End Function
Public Function Create(ParamArray args() As Variant) As List

    Set Create = Copy(CVar(args))
    
End Function
Public Function Repeat(ByVal x, ByVal n As Long) As List

    Set Repeat = defRepeat.Repeat(List, x, n)

End Function
'
' ### Functional Constructors
'
Public Function Map(ByVal op As IApplicable) As List

    Set Map = defMap.Map(List, op, Me)

End Function
Public Function FlatMap(ByVal op As IApplicable) As List

    Set FlatMap = defMap.FlatMap(List, op, Me)

End Function
Public Function Filter(ByVal pred As IApplicable) As List
    
    Set Filter = defFilter.Filter(List, pred, Me)
    
End Function
Public Function FilterNot(ByVal pred As IApplicable) As List
    
    Set FilterNot = defFilter.FilterNot(List, pred, Me)
    
End Function
Public Function Fold(ByVal init As Variant, ByVal binaryOp As IApplicable) As Variant
    
    Assign Fold, defFold.Fold(binaryOp, init, Me)
    
End Function
Public Function Reduce(ByVal binaryOp As IApplicable) As Variant
    
    cast.Assign Reduce, Slice(2, -1).Fold(pCollec(1), binaryOp)
    
End Function
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
' General pattern is to have a property or method that applies to one
' element and then extend that to sequences.
'
' ### Access and Replace
'
' `Item` allows for access and replacement of an element through indexes.
' `Slice` extends that to sequences that allows reversing and steps. Only
' `Slice has steps, other sequence methods only apply to continuous sequences.
'
''
' Default member is Item
Public Property Let Item(ByVal index As Long, ByVal element As Variant)
Attribute Item.VB_UserMemId = 0
    
    TransformIndex index
    Replace index, element
    
End Property
Public Property Set Item(ByVal index As Long, ByVal element As Variant)
Attribute Item.VB_UserMemId = 0
    
    TransformIndex index
    Replace index, element
    
End Property
Public Property Get Item(ByVal index As Long) As Variant
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
' ### Removal
'
''
' Remove
Public Sub Remove(ByVal index As Long)
    
    TransformIndex index
    pCollec.Remove index
    
End Sub
''
' Revisions from CR
Public Sub Clear(ByVal start As Long, ByVal size As Long)
    
    TransformIndex start
    
    Dim i As Long
    For i = 1 To size
        pCollec.Remove start
    Next i

End Sub

'
' ### Extension
'
''
' Append: Vanilla `Add`
Public Sub Append(ParamArray Elements() As Variant)
    Extend CVar(Elements)
End Sub
''
' Extend: Append sequence
' `myList.Extend rng.Value ' shtbrx`
Public Sub Extend(ByVal sequence As Variant)
    
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
' ### Insert
'
''
' Emplace: Place element at index, existing elements shifted backward.
' Synonymous with `pCollec.Add element, before:=index.`
Public Sub Emplace(ByVal index As Long, ByVal element As Variant)
    
    TransformIndex index
    pCollec.Add element, Before:=index
    
End Sub
''
' Insert: Sames as Emplace but with a sequence
Public Sub Insert(ByVal index As Long, ByVal sequence As Variant)
    
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
' ### Auxillary
'
Public Function ToArray() As Variant()

    If pCollec.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    Dim a() As Variant
    ReDim a(pCollec.Count - 1) As Variant
    
    Dim i As Long
    For i = 1 To pCollec.Count
        cast.Assign a(i - 1), pCollec(i)
    Next i
    
    ToArray = a
    
End Function
Public Function Reverse() As List
    
    Set Reverse = Slice(-1, 1)
    
End Function
Public Function Count() As Long

    Count = pCollec.Count
    
End Function
'
' Interfaces
' ----------
'
' ### IBuildable
'
Private Function IBuildable_MakeEmpty() As IBuildable

    Set IBuildable_MakeEmpty = List.Create()

End Function
Private Sub IBuildable_AddItem(ByVal element)

    Append element

End Sub
Private Sub IBuildable_AddItems(ByVal sequence)

    Extend sequence

End Sub
'
' ### IIterable
'
Private Function IIterable_Item(ByVal index As Long) As Variant

    Assign IIterable_Item, Item(index)

End Function
Private Function IIterable_LowerBound() As Long

    IIterable_LowerBound = 1

End Function
Private Function IIterable_UpperBound() As Long

    IIterable_UpperBound = Count

End Function
'
' ### IEquatable
'
Public Function Equals(ByVal x) As Boolean

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
Private Function IEquatable_Equals(ByVal other) As Boolean

    IEquatable_Equals = Equals(other)
    
End Function
'
' ### IMonadic
'
Private Function IMonadic_Bind(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Bind = FlatMap(op)

End Function
Private Function IMonadic_Map(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Map = Map(op)

End Function
Private Function IMonadic_Unit(ByVal x) As IMonadic

    Set IMonadic_Unit = List.Create(x)

End Function
'
'  ### IShowable
'
Public Function Show() As String

    Show = defShow.ShowableObject(Me, ToArray)
    
End Function



Private Function IShowable_Show() As String

    IShowable_Show = Show()
    
End Function
