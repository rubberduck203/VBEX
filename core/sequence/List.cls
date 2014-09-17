VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' List Class
' ==========
'
' Base Linear Container Class
'
' Copywrite (C) 2014 Philip Wales
' This file (List.cls) is distributed under the MIT license
'
' Decisions:
'
'  - Error Handling
'    + Bubble up errors wherever they would arrise from Collection
'    + Keep methods simple enough that bubbling is unnecessary
'
Option Explicit
'
' Private Members
' ---------------
'
Private collec As Collection ' Sole datamember
'
' Public Members
' --------------
'
' None
'
' Private Methods
' ---------------
''
' TransformIndex: Eforces Zero Offset and `Cylcing?.
Private Sub TransformIndex(ByRef x As Variant)
    If x < 0 Then x = x + collec.Count + 1
End Sub
''
' Replace
Private Sub Replace(ByVal index As Long, ByVal element As Variant)
    
    collec.Remove index
    
    If index = collec.Count + 1 Then
        collec.Add element
    Else
        collec.Add element, before:=index
    End If
    
End Sub
'
''
' Constructor
Private Sub Class_Initialize()
    Set collec = New Collection
End Sub
''
' Destructor: Might be unnecessary
Private Sub Class_Terminate()
    Set collec = Nothing
End Sub
''
' Enumerator: How we all would love to create our own...
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    
    Set NewEnum = collec.[_NewEnum]
    
End Property
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
    seq.Assign Item, collec.Item(index)
    
End Property
Public Property Get Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1) As List
    
    TransformIndex a
    TransformIndex b
    
    Set Slice = New List
    
    If s < 1 Then Err.Raise "List.Slice", "Step " & s & " is not a natural number."
    s = IIF(a < b, s, -s)
    
    Dim i As Long
    For i = a To b Step s
        Slice.Append collec.Item(i)
    Next i
    
End Property
Public Property Let Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1, ByVal sequence As Variant)
    
    TransformIndex a
    TransformIndex b
    
    If s < 1 Then Err.Raise "List.Slice", "Step " & s & " is not a natural number."
    s = IIF(a < b, s, -s)
    
    If Abs(a - b) + 1 <> seq.Length(sequence) Then
        Err.Raise 9, "List.Slice", "Subscript out of Range."
    End If
    
    Dim i As Long: i = a
    Dim element As Variant
    For Each element In sequence
        
        Replace i, element
        i = i + s
        
    Next element
    
    Debug.Assert (i - s = b)
    
End Property
'
' ### Removal
'
''
' Remove
Public Sub Remove(ByVal index As Long)
    
    TransformIndex index
    collec.Remove index
    
End Sub
''
' Revisions from CR
Public Sub Clear(ByVal start As Long, ByVal count As Long)
    
    TransformIndex start
    
    Dim i As Long
    For i = 0 To count
        collec.Remove start
    Next i

End Sub

'
' ### Extension
'
''
' Append: Vanilla `Add`
Public Sub Append(ByVal element As Variant)
    collec.Add element
End Sub
''
' Extend: Append sequence
' `myList.Extend rng.Value ' shtbrx`
Public Sub Extend(ByVal sequence As Variant)
    
    Dim element As Variant
    For Each element In sequence
        collec.Add element
    Next element
    
End Sub
'
' ### Insert
'
''
' Emplace: Place element at index, existing elements shifted backward.
' Synonymous with `collec.Add element, before:=index.`
Public Sub Emplace(ByVal index As Long, ByVal element As Variant)
    
    TransformIndex index
    collec.Add element, before:=index
    
End Sub
''
' Insert: Sames as Emplace but with a sequence
Public Sub Insert(ByVal index As Long, ByVal sequence As Variant)
    
    TransformIndex index
    seq.Reverse sequence ' sequence becomes a stack
    
    Dim element As Variant
    For Each element In sequence
        ' push back previous element
        collec.Add element, before:=index
    Next element
    
End Sub
'
' ### Auxillary
'
''
' Count
Public Property Get Count() As Long
    Count = collec.Count
End Property
''
' Exists: Linear search for `sought`.
Public Function Exists(ByVal sought As Variant) As Boolean
    
    Exists = True
    
    Dim element As Variant
    For Each element In collec
        If element = sought Then Exit Function
    Next element
    
    Exists = False
    
End Function
''
' ToString: Represent instance as String.
' Usefull for debugging only.  Doesn't work with nested
' lists.
Public Property Get ToString() As String
    ToString = "[" & Join(seq.ToArray(collec), ", ") & "]"
End Property
