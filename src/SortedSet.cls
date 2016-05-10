VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' SortedSet
' =========
'
' Maintains a list of sorted unique items.
' It is meant to be synonymous with Python's `set` but that is
' a reserved keyword in VBA.  Instead of using an advanced data
' structure, `SortedSet` keeps a sorted `List` and an external
' binary search function to preserve unique elements.  This allows
' _mutability_ and limits recursion (as VB is not tail recursive.)
'
' Copywrite (C) 2014 Philip Wales
' This file (SortedSet.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Option Explicit

Implements Equatable
Implements Showable
Implements SetLike
Implements Buildable
'
' Public Members
' --------------
' None
'
' Private Memebers
' ----------------
'
Private pList As List
'
' Private Methods
' ---------------
'
''
' IndexOf: Finds where an item is or where it should be
Private Function IndexOf(ByVal element) As Long

    If pList.Count > 0 Then
        On Error GoTo ErrHandler
        IndexOf = srch.BinarySearch(element, pList, nearest:=True).GetItem
        On Error GoTo 0
    Else
        IndexOf = 1
    End If
Exit Function
ErrHandler:
    Select Case Err.Number
    Case Else
        Exceptions.BubbleError Me, "IndexOf", Err
    End Select
End Function
''
' Destructor, might be unnecessary
Private Sub Class_Initialize()

    Set pList = New List
    
End Sub
Private Sub Class_Terminate()

    Set pList = Nothing
    
End Sub
'
' Public Methods
' --------------
'
''
' Enumerator: How we all would love to create our own...
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    
    Set NewEnum = pList.NewEnum
    
End Property
'
' Constructors
' ------------
'
Public Function Create(ParamArray args() As Variant) As SortedSet
    
    Set Create = Copy(args)
    
End Function
Public Function Copy(ByVal sequence) As SortedSet
    
    Dim result As New SortedSet
    result.Extend sequence
    Set Copy = result
    
End Function
'
' ### Functional Constructors
'
Public Function Map(ByVal op As Applicable) As SortedSet
    
    Set Map = defMap.TransversableMap(SortedSet, op, pList)
    
End Function
Public Function FlatMap(ByVal op As Applicable) As SortedSet
    
    Set FlatMap = defMap.TransversableBind(SortedSet, op, pList)
    
End Function
Public Function Filter(ByVal pred As Applicable) As SortedSet
    
    Set Filter = defFilter.Filter(SortedSet, pred, pList)
    
End Function
Public Function FilterNot(ByVal pred As Applicable) As SortedSet
    
    Set FilterNot = defFilter.FilterNot(SortedSet, pred, pList)
    
End Function
'
' ### Default Methods Implemented
'
' #### Access
'
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    
    cast.Assign Item, pList.Item(index)
    
End Property
'
' Mutability
' ----------
'
' #### Extension
''
' Add: Instead of Insert, Emplace, or Append
Public Sub Add(ByVal element)
    
    If pList.Count = 0 Then
    
        pList.Append element
        Exit Sub
        
    End If
    
    Dim index As Long
    index = IndexOf(element)
    
    Dim curVal
    Assign curVal, pList(index)
    
    If Not defEquals.Equals(curVal, element) Then
    
        If defCompare.LessThan(curVal, element) Then  ' ONLY HAPPENS AT END?
            pList.Append element
        Else
            pList.Emplace index, element
        End If
        
    End If
    
End Sub
''
' Extend: routed through our `Add`
Public Sub Extend(ByVal sequence)
    
    Dim element
    For Each element In sequence
        Add element
    Next
    
End Sub
'
' #### Removal
'
''
' Removal is by element not by index
Public Sub Remove(ByVal element)
    
    Dim index As Long
    index = IndexOf(element)
    
    If pList(index) = element Then
        pList.Remove index
    Else
    
        Dim msg As String
        msg = defShow.Show(element) & " does not exist."
        IndexError Me, "Remove", msg
        
    End If
    
End Sub
Public Sub Discard(ByVal element As Variant)

    Dim index As Long
    index = IndexOf(element)
    
    If defEquals.Equals(pList(index), element) Then
        pList.Remove index
    End If
    
End Sub
Public Sub Clear()

    pList.Clear 1, pList.Count
    
End Sub
'
' Interfaces
' ----------
'
'
' ### Buildable
'
Private Function Buildable_MakeEmpty() As Buildable

    Set Buildable_MakeEmpty = Create()

End Function
Private Sub Buildable_AddItem(ByVal element)

    Add element

End Sub
Private Sub Buildable_AddItems(ByVal sequence)

    Extend sequence
    
End Sub
'
' ### SetLike
'
Public Function Contains(ByVal sought) As Boolean

    If pList.Count > 0 Then
    
        Dim index As Long
        index = IndexOf(sought)
        
        Contains = defEquals.Equals(pList.Item(index), sought)
    Else
        Contains = False
    End If
    
End Function
Public Function Count() As Long

    Count = pList.Count

End Function
Public Function IsDisJoint(ByVal other As SetLike) As Boolean

    IsDisJoint = defSetLike.IsDisJoint(Me, other)
    
End Function
Public Function IsSubSetOf(ByVal other As SetLike) As Boolean
    
    IsSubSetOf = defSetLike.IsSubSetOf(Me, other)
    
End Function
Public Function IsProperSubSetOf(ByVal other As SetLike) As Boolean

    IsProperSubSetOf = defSetLike.IsProperSubSetOf(Me, other)
    
End Function
Public Function IsSuperSetOf(ByVal other As SetLike) As Boolean

    IsSuperSetOf = defSetLike.IsSuperSetOf(Me, other)
    
End Function
Public Function IsProperSuperSetOf(ByVal other As SetLike) As Boolean

    IsProperSuperSetOf = defSetLike.IsProperSuperSetOf(Me, other)
    
End Function
Public Function Union(ByVal other As SetLike) As SortedSet

    Set Union = defSetLike.Union(SortedSet, Me, other)
    
End Function
Public Function Intersect(ByVal other As SetLike) As SortedSet

    Set Intersect = defSetLike.Intersect(SortedSet, Me, other)

End Function
Public Function Difference(ByVal other As SetLike) As SortedSet
    
    Set Difference = defSetLike.Difference(SortedSet, Me, other)
    
End Function
Public Function SymmetricDifference(ByVal other As SetLike) As SortedSet

    Set SymmetricDifference = defSetLike.SymmetricDifference(SortedSet, Me, other)
    
End Function
Private Property Get SetLike_Elements() As IUnknown
    
    Set SetLike_Elements = Me.NewEnum
    
End Property
Private Function SetLike_Contains(ByVal sought) As Boolean

    SetLike_Contains = Contains(sought)

End Function
Private Function SetLike_Count() As Long

    SetLike_Count = Count

End Function
Private Function SetLike_IsDisJoint(ByVal other As SetLike) As Boolean

    SetLike_IsDisJoint = IsDisJoint(other)

End Function
Private Function SetLike_IsSubSetOf(ByVal other As SetLike) As Boolean

    SetLike_IsSubSetOf = IsSubSetOf(other)

End Function
Private Function SetLike_IsProperSubSetOf(ByVal other As SetLike) As Boolean

    SetLike_IsProperSubSetOf = IsProperSubSetOf(other)

End Function
Private Function SetLike_IsSuperSetOf(ByVal other As SetLike) As Boolean

    SetLike_IsSuperSetOf = IsSuperSetOf(other)

End Function
Private Function SetLike_IsProperSuperSetOf(ByVal other As SetLike) As Boolean

    SetLike_IsProperSuperSetOf = IsProperSuperSetOf(other)

End Function
Private Function SetLike_Union(ByVal other As SetLike) As SetLike
    
    Set SetLike_Union = Union(other)

End Function
Private Function SetLike_Intersect(ByVal other As SetLike) As SetLike

    Set SetLike_Intersect = Intersect(other)

End Function
Private Function SetLike_Difference(ByVal other As SetLike) As SetLike

    Set SetLike_Difference = Difference(other)

End Function
Private Function SetLike_SymmetricDifference(ByVal other As SetLike) As SetLike

    Set SetLike_SymmetricDifference = SymmetricDifference(other)

End Function
'
' ### Equatable
'
Public Function Equals(ByVal other) As Boolean

    Equals = defSetLike.SetEquals(Me, other)
    
End Function
Private Function Equatable_Equals(ByVal other) As Boolean

    Equatable_Equals = Equals(other)
    
End Function
'
' ### Showable
'
Public Function Show() As String

    Show = defShow.ShowableObject(Me, pList.ToArray)
    
End Function
Private Function Showable_Show() As String

    Showable_Show = Show
    
End Function
