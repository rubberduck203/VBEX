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
Public Function Create(ParamArray tuples() As Variant) As Dict

    Set Create = FromTuples(List.Copy(tuples))
    
End Function
Public Function FromLists(ByVal keyList As List, ByVal valueList As List) As Dict
    Debug.Assert (keyList.Count = valueList.Count) ' TODO: raise error
    
    Dim result As New Dict
    
    ' TODO: use Zip. then expect users to use `FromTuples`?
    Dim i As Long
    For i = 1 To keyList.Count
        result.Add keyList(i), valueList(i)
    Next i
    
    Set FromLists = result
    
End Function
' TODO: use a defined type not a tuple but the type must be public....
Public Function FromTuples(ByVal tuples As List) As Dict
    
    Dim result As New Dict
    Dim keyValueTuple As Variant
    For Each keyValueTuple In tuples
        
        result.Add keyValueTuple.Item(1), keyValueTuple.Item(2)
        
    Next keyValueTuple
    
    Set FromTuples = result
    
End Function

Public Function Copy(ByVal thatD As Dict) As Dict
    Set Copy = FromTuples(thatD.Pairs)
End Function
'
' Public Properties
' -----------------
'
' ### Single Item actions
'
''
' `Item`: same as usual
' TODO: how does pDict do this?
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    cast.Assign Item, pDict(key)
End Property
Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
Attribute Item.VB_UserMemId = 0
    pDict(key) = value
End Property
Public Property Set Item(ByVal key As Variant, ByVal value As Variant)
Attribute Item.VB_UserMemId = 0
    Set pDict(key) = value
End Property
''
' GetItem: Return default value if `key` does not exist
Public Property Get GetItem(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Variant
    If pDict.Exists(key) Then
        cast.Assign GetItem, pDict(key)
    Else
        cast.Assign GetItem, default
    End If
End Property
''
' Pop: return value, and remove value, has default implementation
Public Property Get Pop(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Variant
    If pDict.Exists(key) Then
        cast.Assign Pop, pDict(key)
        pDict.Remove key
    Else
        cast.Assign Pop, default
    End If
End Property
''
' Pair: return a `Tuple` of (key, value)
Public Property Get Pair(ByVal key As Variant) As Tuple
    Set Pair = Tuple.Pack(key, pDict(key))
End Property
''
' PopPair: return Pair and remove them.
Public Property Get PopPair(ByVal key As Variant, _
        Optional ByVal default As Variant = Empty) As Tuple
    If pDict.Exists(key) Then
        Set PopPair = Pair(key)
        pDict.Remove key
    Else
        Set PopPair = Tuple.Pack(key, default)
    End If
End Property
'
' Produce Lists
' -------------
'
Public Property Get Keys() As List
    Set Keys = List.Copy(pDict.Keys)
End Property
Public Property Get Values() As List
    Set Values = List.Copy(pDict.Items)
End Property
Public Property Get Pairs() As List
    Set Pairs = Tuple.Zip(Keys(), Values())
End Property
Public Property Get Count() As Long
    Count = pDict.Count
End Property
'
' Public Methods
' --------------
'
''
' `Add`: add item to dictionary TODO: Raise error if already exists
Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    pDict.Add key, value
End Sub
''
' Remove: Remove entry
Public Sub Remove(ByVal key As Variant)
    pDict.Remove key
End Sub
Public Function Exists(ByVal key As Variant) As Boolean
    Exists = pDict.Exists(key)
End Function
'
''
' Update: add another `Dict` to this `Dict` overwriting when
' a key alread exists
Public Sub Update(ByVal other As Dict)
    Dim key As Variant
    For Each key In other.Keys
        pDict(key) = other(key)
    Next key
End Sub
Public Sub Clear(ByVal key As Variant)
    pDict.RemoveAll
End Sub

Public Property Get ToString(Optional ByVal delim As String = ", ") As String

    Dim result As String
    
    Dim key As Variant
    For Each key In pDict.Keys()
        result = result & delim & key & ": " & pDict.Item(key)
    Next key
    
    ToString = "{" & Mid(result, Len(delim) + 1) & "}"
    
End Property
