VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' HashSet
' =======
'
Option Explicit

Implements IEquatable
Implements IShowable
Implements ISetLike
Implements IBuildable
'
' Public Members
' --------------
' None
'
' Private Memebers
' ----------------
'
Private pDict As Scripting.Dictionary
'
' Private Methods
' ---------------
'
''
' Destructor, might be unnecessary
Private Sub Class_Initialize()

    Set pDict = New Scripting.Dictionary
    
End Sub
Private Sub Class_Terminate()

    Set pDict = Nothing
    
End Sub
'
' Public Methods
' --------------
'
''
' Enumerator: How we all would love to create our own...
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    
    Set NewEnum = pDict.Keys
    
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
Public Function Map(ByVal op As IApplicable) As SortedSet
    
    Set Map = defMap.Map(SortedSet, op, pDict)
    
End Function
Public Function FlatMap(ByVal op As IApplicable) As SortedSet
    
    Set FlatMap = defMap.FlatMap(SortedSet, op, pDict)
    
End Function
Public Function Filter(ByVal pred As IApplicable) As SortedSet
    
    Set Filter = defFilter.Filter(SortedSet, pred, pDict)
    
End Function
Public Function FilterNot(ByVal pred As IApplicable) As SortedSet
    
    Set FilterNot = defFilter.FilterNot(SortedSet, pred, pDict)
    
End Function
'
' ### Default Methods Implemented
'
' #### Access
'
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    
    cast.Assign Item, pDict.Item(index)
    
End Property
'
' Mutability
' ----------
'
' #### Extension
''
' Add: Instead of Insert, Emplace, or Append
Public Sub Add(ByVal element)
    
    If Contains(element) Then
        pDict.Add element, Nothing
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
    
    If Contains(element) Then
        pDict.Remove element
    Else
    
        Dim msg As String
        msg = defShow.Show(element) & ", does not exist."
        IndexError Me, "Remove", msg
        
    End If
    
End Sub
Public Sub Discard(ByVal element)

    pDict.Remove element
    
End Sub
Public Sub Clear()

    pDict.Clear
    
End Sub
'
' Interfaces
' ----------
'
'
' ### IBuildable
'
Private Function IBuildable_MakeEmpty() As IBuildable

    Set IBuildable_MakeEmpty = Create()

End Function
Private Sub IBuildable_AddItem(ByVal element)

    Add element

End Sub
Private Sub IBuildable_AddItems(ByVal sequence)

    Extend sequence
    
End Sub
'
' ### ISetLike
'
Public Function Contains(ByVal sought) As Boolean

    Contains = pDict.Exists(sought)
    
End Function
Public Function Count() As Long

    Count = pDict.Count

End Function
Public Function IsDisJoint(ByVal other As ISetLike) As Boolean

    IsDisJoint = defSetLike.IsDisJoint(Me, other)
    
End Function
Public Function IsSubSetOf(ByVal other As ISetLike) As Boolean
    
    IsSubSetOf = defSetLike.IsSubSetOf(Me, other)
    
End Function
Public Function IsProperSubSetOf(ByVal other As ISetLike) As Boolean

    IsProperSubSetOf = defSetLike.IsProperSubSetOf(Me, other)
    
End Function
Public Function IsSuperSetOf(ByVal other As ISetLike) As Boolean

    IsSuperSetOf = defSetLike.IsSuperSetOf(Me, other)
    
End Function
Public Function IsProperSuperSetOf(ByVal other As ISetLike) As Boolean

    IsProperSuperSetOf = defSetLike.IsProperSuperSetOf(Me, other)
    
End Function
Public Function Union(ByVal other As ISetLike) As SortedSet

    Set Union = defSetLike.Union(SortedSet, Me, other)
    
End Function
Public Function Intersect(ByVal other As ISetLike) As SortedSet

    Set Intersect = defSetLike.Intersect(SortedSet, Me, other)

End Function
Public Function Difference(ByVal other As ISetLike) As SortedSet
    
    Set Difference = defSetLike.Difference(SortedSet, Me, other)
    
End Function
Public Function SymmetricDifference(ByVal other As ISetLike) As SortedSet

    Set SymmetricDifference = defSetLike.SymmetricDifference(SortedSet, Me, other)
    
End Function
Private Property Get ISetLike_Elements() As IUnknown
    
    Set ISetLike_Elements = Me.NewEnum
    
End Property
Private Function ISetLike_Contains(ByVal sought) As Boolean

    ISetLike_Contains = Contains(sought)

End Function
Private Function ISetLike_Count() As Long

    ISetLike_Count = Count

End Function
Private Function ISetLike_IsDisJoint(ByVal other As ISetLike) As Boolean

    ISetLike_IsDisJoint = IsDisJoint(other)

End Function
Private Function ISetLike_IsSubSetOf(ByVal other As ISetLike) As Boolean

    ISetLike_IsSubSetOf = IsSubSetOf(other)

End Function
Private Function ISetLike_IsProperSubSetOf(ByVal other As ISetLike) As Boolean

    ISetLike_IsProperSubSetOf = IsProperSubSetOf(other)

End Function
Private Function ISetLike_IsSuperSetOf(ByVal other As ISetLike) As Boolean

    ISetLike_IsSuperSetOf = IsSuperSetOf(other)

End Function
Private Function ISetLike_IsProperSuperSetOf(ByVal other As ISetLike) As Boolean

    ISetLike_IsProperSuperSetOf = IsProperSuperSetOf(other)

End Function
Private Function ISetLike_Union(ByVal other As ISetLike) As ISetLike
    
    Set ISetLike_Union = Union(other)

End Function
Private Function ISetLike_Intersect(ByVal other As ISetLike) As ISetLike

    Set ISetLike_Intersect = Intersect(other)

End Function
Private Function ISetLike_Difference(ByVal other As ISetLike) As ISetLike

    Set ISetLike_Difference = Difference(other)

End Function
Private Function ISetLike_SymmetricDifference(ByVal other As ISetLike) As ISetLike

    Set ISetLike_SymmetricDifference = SymmetricDifference(other)

End Function
'
' ### IEquatable
'
Public Function Equals(ByVal other) As Boolean

    Equals = pDict.Equals(other)
    
End Function
Private Function IEquatable_Equals(ByVal other) As Boolean

    IEquatable_Equals = Equals(other)
    
End Function
'
' ### IShowable
'
Public Function Show() As String

    Show = defShow.ShowableObject(Me, List.Copy(pDict.Keys).ToArray)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function
