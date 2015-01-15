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
' This file (SortedSet.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
Option Explicit

Implements IEquatable
Implements ICountable
Implements IPrintable
Implements ISequence

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
Private Function IndexOf(ByVal element As Variant) As Long
On Error GoTo ErrHandler

    If pList.Count > 0 Then
        IndexOf = srch.BinarySearch(element, pList, 1, pList.Count, nearest:=True)
    Else
        IndexOf = 1
    End If
    
Exit Function
ErrHandler:
    Debug.Assert Err.number = 438
    RaiseObjectWithNoDefaultPropertyError
    
End Function
Private Sub RaiseObjectWithNoDefaultPropertyError()

    Err.Raise 438, "SortedSet.IndexOf", _
        "Object does not have a default property and cannot be compared"
        
End Sub
''
' Destructor, might be unnecessary
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
    
    Set NewEnum = pList.NewEnum
    
End Property
'
' Constructors
' ------------
'
Private Sub Class_Initialize()

    Set pList = New List
    
End Sub
'
' ### Initializers
'
Public Function Copy(ByVal sequence As Variant) As SortedSet
    
    Dim result As New SortedSet
    result.Extend sequence
    Set Copy = result
    
End Function
Public Function Create(ParamArray args() As Variant) As SortedSet
    
    Set Create = Copy(args)
    
End Function
'
' Immutable Constructors
' ----------------------
'
Public Function Union(ByVal other As SortedSet) As SortedSet

    Set Union = New SortedSet
    Union.Extend other
    Union.Extend Me
    
End Function
Public Function Intersect(ByVal other As SortedSet) As SortedSet
    Dim result As New SortedSet
    
    Dim element As Variant
    For Each element In pList
        If other.Exists(element) Then result.Add element
    Next element
    
    Set Intersect = result
    
    
End Function
Public Function Difference(ByVal other As SortedSet) As SortedSet
    
    Dim result As New SortedSet
    
    Dim element As Variant
    For Each element In pList
        If Not other.Exists(element) Then result.Add element
    Next element
    
    Set Difference = result
    
End Function
Public Function SymmetricDifference(ByVal other As SortedSet) As SortedSet

    Set SymmetricDifference = Difference(other).Union(other.Difference(Me))
    
End Function
'
' ### Functional Constructors
'
Public Function Map(ByVal op As IApplicable) As SortedSet
    
    Set Map = Copy(pList.Map(op))
    
End Function
Public Function Filter(ByVal pred As IApplicable, _
        Optional ByVal opposite As Boolean = False) As SortedSet
    
    Set Filter = Copy(pList.Filter(pred, opposite))
    
End Function
Public Function ForAll(ByVal pred As IApplicable, _
        Optional ByVal opposite As Boolean = False) As Boolean

    ForAll = pList.ForAll(pred, opposite)

End Function
'
'
'
' Comparision
' -----------
'
''
' Exists: Use Binary Search for faster results.
Public Function Exists(ByVal sought As Variant) As Boolean

    If pList.Count > 0 Then
        Exists = (pList(IndexOf(sought)) = sought)
    Else
        Exists = False
    End If
    
End Function
Public Function IsDisJoint(ByVal other As SortedSet) As Boolean

    IsDisJoint = (Intersect(other).Count = 0)
    
End Function
Public Function IsSubSetOf(ByVal other As SortedSet) As Boolean
    
    Dim element As Variant
    For Each element In pList
        If Not other.Exists(element) Then
            IsSubSetOf = False
            Exit Function
        End If
    Next element
    
    IsSubSetOf = True
    
End Function
Public Function IsProperSubSetOf(ByVal other As SortedSet) As Boolean

    IsProperSubSetOf = (IsSubSetOf(other) And (Count < other.Count))
    
End Function
Public Function IsSuperSetOf(ByVal other As SortedSet)

    IsSuperSetOf = other.IsSubSetOf(Me)
    
End Function
Public Function IsProperSuperSetOf(ByVal other As SortedSet) As Boolean

    IsProperSuperSetOf = other.IsProperSubSetOf(Me)
    
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
' ### Overridden Methods
'
' #### Extension
''
' Add: Instead of Insert, Emplace, or Append
Public Sub Add(ByVal element As Variant)
    
    If pList.Count = 0 Then
        pList.Append element
        Exit Sub
    End If
    
    Dim index As Long
    index = IndexOf(element)
    
    If pList(index) = element Then Exit Sub
    
    If pList(index) < element Then ' ONLY HAPPENS AT END?
        pList.Append (element)
    Else
        pList.Emplace index, element
    End If
    
End Sub
''
' Extend: routed through our `Add`
Public Sub Extend(ByVal sequence As Variant)
    
    Dim element As Variant
    For Each element In sequence
        Add element
    Next element
    
End Sub
'
' #### Removal
'
''
' Removal is by element not by index
Public Sub Remove(ByVal element As Variant)
    
    Dim index As Long
    index = IndexOf(element)
    
    If pList(index) = element Then
        pList.Remove index
    Else
        Err.Raise 9, TypeName(Me) & ".Remove", "Element does not exist."
    End If
    
End Sub
Public Sub Discard(ByVal element As Variant)

    Dim index As Long
    index = IndexOf(element)
    
    If pList(index) = element Then pList.Remove index
    
End Sub
Public Sub Clear()

    pList.Clear 1, pList.Count
    
End Sub
Public Function Pop() As Variant

    cast.Assign Pop, pList(1)
    pList.Remove 1
    
End Function
'
' Interfaces
' ----------
'
' ### ISequence
'
Public Function ToArray() As Variant()

    ToArray = pList.ToArray()
    
End Function
Public Function ISequence_ToArray() As Variant()
  
    ISequence_ToArray = ToArray
  
End Function
'
' ### IEquatable
'
Public Function Equals(ByVal other As Variant) As Boolean

    Equals = pList.Equals(other)
    
End Function
Public Function IEquatable_Equals(ByVal other As Variant) As Boolean

    IEquatable_Equals = Equals(other)
    
End Function
'
' ### ICountable
'
Public Property Get Count() As Long

    Count = pList.Count
    
End Property
Public Property Get ICountable_Count() As Long

    ICountable_Count = Count
    
End Property
'
' ### IPrintable
'
Public Property Get ToString(Optional delim As String = ", ") As String

    ToString = cast.SequenceToString(Me, delim, lcap:="{", rcap:="}")
    
End Property
Public Property Get IPrintable_ToString() As String

    IPrintable_ToString = ToString
    
End Property
