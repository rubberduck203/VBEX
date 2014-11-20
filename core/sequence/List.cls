VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' List Class
' ==========
'
' Base Linear Container Class.  Based on Python List class.
'
' Copywrite (C) 2014 Philip Wales
' This file (List.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
'
Option Explicit
Implements IEquatable
Implements ICountable
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
''
' TransformIndex: Eforces Zero Offset and `Cylcing?.
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
        pCollec.Add element, before:=index
    End If
    
End Sub
Private Sub RaiseNotSequenceError(ByVal method As String, ByVal scalar As Variant)
    Err.Raise 13, TypeName(Me) & "." & method, """" & TypeName(scalar) & """ is not an iterable type."
End Sub
'
''
' Constructor
Public Function Copy(ByVal sequence As Variant) As List

    Dim result As New List
    result.Extend sequence
    Set Copy = result
    
End Function
Public Function Create(ParamArray args() As Variant) As List

    Set Create = Copy(args)
    
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
    seq.Assign Item, pCollec.Item(index)
    
End Property
Public Property Get Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1) As List
    
    TransformIndex a
    TransformIndex b
    
    Set Slice = New List
    
    If s < 1 Then Err.Raise 9, TypeName(Me) & ".Slice", "Step " & s & " is not a natural number."
    s = IIF(a < b, s, -s)
    
    Dim i As Long
    For i = a To b Step s
        Slice.Append pCollec.Item(i)
    Next i
    
End Property
Public Property Let Slice(ByVal a As Long, ByVal b As Long, Optional ByVal s As Long = 1, ByVal sequence As Variant)
    
    TransformIndex a
    TransformIndex b
    
    If s < 1 Then Err.Raise 9, TypeName(Me) & ".Slice", "Step " & s & " is not a natural number."
    s = IIF(a < b, s, -s)
    
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
Public Sub Append(ParamArray elements() As Variant)
    Extend elements
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
    pCollec.Add element, before:=index
    
End Sub
''
' Insert: Sames as Emplace but with a sequence
Public Sub Insert(ByVal index As Long, ByVal sequence As Variant)
    
    TransformIndex index
    On Error GoTo NotSequence
    
    Dim other As List
    Set other = List.Copy(sequence)
    sort.Reverse other, 1, other.Count
    
    Dim element As Variant
    For Each element In other
        ' push back previous element
        pCollec.Add element, before:=index
    Next element

Exit Sub
NotSequence:
    RaiseNotSequenceError "Insert", sequence
    
End Sub
'
' ### Auxillary
'
''
' Count
Public Property Get Count() As Long
    Count = pCollec.Count
End Property
''
' Exists: Linear search for `sought`.
Public Function Exists(ByVal sought As Variant) As Boolean
    
    Exists = True
    
    Dim element As Variant
    For Each element In pCollec
        If element = sought Then Exit Function
    Next element
    
    Exists = False
    
End Function
Public Function Equals(ByVal other As Variant) As Boolean
    On Error GoTo Nope
    
    Equals = False
    If other.Count <> Count Then Exit Function
    
    Dim i As Long
    For i = 1 To Count
        If other(i) <> pCollec(i) Then Exit Function
    Next i
    
    Equals = True
    
CleanExit:
Exit Function
Nope:

    Equals = False
    Resume CleanExit
    
End Function
''
' ToString: Represent instance as String.
' Usefull for debugging only.  Doesn't work with nested
' lists.
Public Function ToString(Optional delim As String = ", ") As String
    ToString = "[" & Join(ToArray, delim) & "]"
End Function
Public Function ToArray() As Variant()

    If pCollec.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    Dim a() As Variant
    ReDim a(pCollec.Count - 1) As Variant
    Dim i As Long
    For i = 1 To pCollec.Count
        a(i - 1) = pCollec(i)
    Next i
    ToArray = a
End Function
'
' Interfaces
' ----------
'
' ### IEquatable
'
Public Function IEquatable_Equals(ByVal other As Variant) As Boolean
    IEquatable_Equals = Equals(other)
End Function
'
' ### ICountable
'
Public Property Get ICountable_Count() As Long
    ICountable_Count = Count
End Property
