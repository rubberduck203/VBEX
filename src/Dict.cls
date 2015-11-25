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
Implements IShowable
' Should implement more Interfaces but this is the only mapping type yet.
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
Public Function Create(ParamArray assocs()) As Dict

    Set Create = FromAssocs(List.Copy(assocs))
    
End Function
Public Function FromLists(ByVal ks As IIterable, ByVal vs As IIterable) As Dict
    Debug.Assert (ks.UpperBound = vs.UpperBound And _
        ks.LowerBound = vs.LowerBound) ' TODO: raise error
    
    Dim result As New Dict
    
    ' TODO: use Zip. then expect users to use `FromAssocs`?
    Dim i As Long
    For i = ks.LowerBound To ks.UpperBound
        result.Add ks.Item(i), vs.Item(i)
    Next i
    
    Set FromLists = result
    
End Function
' TODO: use a defined type not a tuple but the type must be public....
Public Function FromAssocs(ByVal assocs) As Dict
    
    Dim result As New Dict
    
    Dim kv
    For Each kv In assocs
        result.Add kv.key, kv.Value
    Next
    
    Set FromAssocs = result
    
End Function
Public Function Copy(ByVal thatD As Dict) As Dict

    Set Copy = FromAssocs(thatD.Pairs)
    
End Function
'
' Public Properties
' -----------------
'
' ### Single Item actions
'
''
' `Item`: same as usual
' TODO: how does pDict do this -> hashtables
Public Property Get Item(ByVal k) As Variant
Attribute Item.VB_UserMemId = 0

    If pDict.Exists(k) Then
        Assign Item, pDict(k)
    Else
        KeyError Me, "Item", defShow.Show(k) & " is not an entry"
    End If
    
End Property
Public Property Let Item(ByVal k, ByVal v)
Attribute Item.VB_UserMemId = 0

    pDict(k) = v
    
End Property
Public Property Set Item(ByVal k, ByVal v)
Attribute Item.VB_UserMemId = 0

    Set pDict(k) = v
    
End Property
''
' GetItem: Return default value if `key` does not exist
Public Function GetItem(ByVal k, Optional ByVal default = Empty) As Variant
        
    Console.PrintLine "WARNING: Method Dict.GetItem is deprecated."
    Assign GetItem, GetOrElse(k, default)
    
End Function
Public Function GetOrElse(ByVal k, Optional ByVal default = Empty) As Variant

    Assign GetOrElse, MaybeGet(k).GetOrElse(default)

End Function
Public Function MaybeGet(ByVal k) As Maybe

    Set MaybeGet = Maybe.Attempt(ByName.Create(OnArgs.Make("Item", VbGet, Me), k))

End Function
''
' Pair: return a `Tuple` of (key, value)
Public Function Pair(ByVal k) As Assoc

    Set Pair = Assoc.Make(k, pDict(k))
    
End Function
'
' Produce Lists
' -------------
'
Public Function keys() As List

    Set keys = List.Copy(pDict.keys)
    
End Function
Public Function Values() As List

    Set Values = List.Copy(pDict.Items)
    
End Function
Public Function Pairs() As List

    Dim result As List
    Set result = List.Create
    
    Dim k
    For Each k In keys
        result.Append Pair(k)
    Next
    
    Set Pairs = result
    
End Function
'
' Public Methods
' --------------
'
Public Function Count() As Long

    Count = pDict.Count
    
End Function
Public Sub Add(ByVal k, ByVal v)

    pDict.Add k, v
    
End Sub
''
' Remove: Remove entry
Public Sub Remove(ByVal k)

    pDict.Remove k
    
End Sub
Public Function Exists(ByVal k) As Boolean

    Exists = pDict.Exists(k)
    
End Function
Public Sub Update(ByVal other As Dict)

    Dim k
    For Each k In other.keys
        pDict(k) = other(k)
    Next
    
End Sub
Public Sub Clear()

    pDict.RemoveAll
    
End Sub
'
' IShowable
' ---------
'
Public Function Show() As String

    Show = defShow.ShowableObject(Me, Pairs.ToArray)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
