VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Dict
' ====
'
' Mutelable Mapping class based on Python's `Dict`
'
' Copywrite (C) 2014 Philip Wales
' This file (Dict.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
' Public Members
' --------------
' None
'
' Private Members
' ---------------
'
Private pDict As Scripting.Dictionary
'
' Constructors
' ------------
'
Private Sub Class_Initialize()
    Set pDict = New Scripting.Dictionary
End Sub
Public Function Create(ParamArray Pairs() As Variant) As Dict
    Set Create = Copy(Pairs)
End Function
Public Function Copy(ByVal Pairs As Variant) As Dict
    If IsRectangular(Pairs) Then
        Set Copy = FromRectangular(Pairs)
    Else
        Set Copy = FromJagged(Pairs)
    End If
End Function
Private Function IsRectangular(ByVal someSequence As Variant) As Boolean
    On Error GoTo Nope
    Dim useless As Variant
    useless = UBound(someSequence, 2)
    IsRectangular = True
CleanExit:
    Exit Function
Nope:
    IsRectangular = False
    Resume CleanExit
End Function
Private Function FromRectangular(ByVal Pairs As Variant) As Dict
    
    Dim keyIndex As Integer
    keyIndex = LBound(Pairs, 2)
    
    Dim valueIndex As Integer
    valueIndex = UBound(Pairs, 2)
    
    If valueIndex - keyIndex <> 1 Then Err.Raise 666
    
    Dim result As New Dict
    
    Dim i As Long
    For i = LBound(Pairs, 1) To UBound(Pairs, 2)
        result.Add Pairs(i, keyIndex), Pairs(i, valueIndex)
    Next i
    
    Set FromRectangular = result
    
End Function
Private Function FromJagged(ByVal Pairs As Variant) As Dict

    Dim result As New Dict
    
    Dim pair As Variant
    For Each pair In Pairs
    
        Dim keyIndex As Integer
        keyIndex = LBound(pair)
        
        Dim valueIndex As Integer
        valueIndex = UBound(pair)
        
        If valueIndex - keyIndex <> 1 Then Err.Raise 666
        result.Add pair(keyIndex), pair(valueIndex)
        
    Next pair
    
    Set FromJagged = result
    
End Function
Public Sub Add(ByVal key As Variant, ByVal value As String)
    pDict.Add key, value
End Sub
'
' Public Accessors
' -----------------
'
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0

    seq.Assign Item, pDict(key)
End Property
Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
Attribute Item.VB_UserMemId = 0

    If pDict.Exists(key) Then
        pDict.Remove key
        pDict.Add key, value
    Else
        pDict.Add key, value
    End If
End Property
''
' do we bother to check if `value` is an object?
'Public Property Set Item(ByVal key As Variant, ByVal value As Variant)
'    Debug.Print "Warning `Set Dict.Item(key) = value` is synonymous with `Dict.Item(key) = value`"
'    Me.Item(key) = value
'End Property
Public Property Get GetItem(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Variant
    If pDict.Exists(key) Then
        seq.Assign GetItem, pDict(key)
    Else
        seq.Assign GetItem, default
    End If
End Property
Public Property Get Pop(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Variant
    If pDict.Exists(key) Then
        seq.Assign Pop, pDict(key)
        pDict.Remove key
    Else
        seq.Assign Pop, default
    End If
End Property
Public Property Get PopItem(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Variant
    If pDict.Exists(key) Then
        Set PopItem = Tuple.Pack(key, pDict(key))
        pDict.Remove key
    Else
        Set PopItem = Tuple.Pack(key, default)
    End If
End Property
Public Sub Update(ByVal other As Dict)
    Dim key As Variant
    For Each key In other.Keys
        Item(key) = other(key)
    Next key
End Sub
Public Sub Remove(ByVal key As Variant)
    pDict.Remove key
End Sub
Public Sub Clear(ByVal key As Variant)
    pDict.RemoveAll
End Sub
'
' Aux Methods
' -----------
'
Public Property Get Keys() As List
    Set Keys = List.Copy(pDict.Keys)
End Property
Public Property Get Values() As List
    Set Values = List.Copy(pDict.Items)
End Property
Public Property Get Items() As List
    Set Items = List.Create
    Dim key As Variant
    For Each key In pDict.Keys
        Items.Append Tuple.Pack(key, pDict(key))
    Next key
End Property
Public Property Get Count() As Long
    Count = pDict.Count
End Property
Public Function ToString(Optional ByVal delim As String = ", ") As String

    Dim result As String
    
    Dim key As Variant
    For Each key In pDict.Keys()
        result = result & delim & key & ": " & pDict(key)
    Next key
    
    ToString = "{" & Mid(result, Len(delim) + 1) & "}"
    
End Function
