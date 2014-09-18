VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Sorted List
' ===========
'
' Maintains a list of sorted unique items
'
' Copywrite (C) 2014 Philip Wales
' This file (SortedList.cls) is distributed under the MIT license
'
Option Explicit

' Private Memebers
' ----------------
'
Private this as List
'
' Private Methods
' ---------------
'
''
' IndexOf: Finds where an item is or where it should be
Private Function IndexOf(ByVal element As Variant) As Long
    IndexOf = seq.BinarySearch(element, this, nearest:=true)
End Function
''
' Constructor
Private Sub Class_Initialize()
    Set this = New List
End Sub
''
' Destructor, might be unnecessary
Private Sub Class_Terminate()
    Set this = Nothing
End Sub
'
' Public Methods
' --------------
''
' Enumerator: How we all would love to create our own...
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    
    Set NewEnum = collec.[_NewEnum]
    
End Property
'
' ### Overridden Methods
'
' #### Extension
''
' Add: Instead of Insert, Emplace, or Append
Public Sub Add(ByVal element As Variant)
    
    If this.Count = 0 Then
    
        this.Append element
        Exit Sub
        
    End If
    
    Dim index As Long
    index = indexOf(element)
    
    If this(index) = element Then Exit Sub
    
    If this(index) < element Then ' ONLY HAPPENS AT END?
        this.append(element)
    Else
        this.Emplace index, element
    End If
    
End Sub
''
' Extend: routed through our `Add`
Public Sub Extend(ByVal sequence As Variant)
    
    Dim element As Variant
    For Each element in sequence
        Add element
    Next element
    
End Sub
'
' #### Removal
'
''
' Removal is by element not by index
Public Sub Remove(ByVal element as Variant)
    
    Dim index As Long
    index = IndexOf(element)
    
    If this(index) = element Then
        this.Remove(index)
    Else
        Err.Raise 9, TypeName(Me) & ".Remove", "Element does not exist."
    End If
    
End Sub
'
' Clear not implemented!
'
'
' #### Aux
'
' Exists: Use Binary Search for faster results.
'
Public Function Exists(ByVal sought As Variant) As Boolean
    Exists = (this(IndexOf(sought)) = sought)
End Function
'
' ### Default Methods Implemented
'
' #### Access
'
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    
    seq.Assign Item, this.Item(index)
    
End Property
Public Property Get Slice(ByVal a As Long, ByVal b As Long) As SortedList
    Set Slice = this.Slice(a, b)
End Property
'
' #### Aux
'
Public Property Get Count() As Long
    Count = this.Count
End Property
Public Property Get ToString() As String
    ToString = this.ToString
End Property
'
' ### Unimplemented
'
' #### Replacement
' #### Insertion
'
