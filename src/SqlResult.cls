VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' SqlResult
' ---------
'
Implements IShowable
Implements IIterable

Private pRecords As Maybe ' Maybe(Rows)
Private pFieldNames As List ' [Fields]
'
' Constructors
' ------------
'
Public Function Make(ByVal rs As ADODB.RecordSet) As SqlResult

    Dim result As New SqlResult
    Set result.RecordSet = rs
    Set Make = result

End Function
Friend Property Set RecordSet(ByVal rs As ADODB.RecordSet)

    Set pFieldNames = TransversableMap(List, OnObject.Make("Name", VbGet), rs.Fields)
    
    If Not rs.EOF Then
        Set pRecords = Maybe.Some(rs.GetRows)
    Else
        Set pRecords = Maybe.None
    End If

End Property
Public Property Get Records() As Maybe

    Set Records = pRecords

End Property
Public Property Get FieldNames() As List

    Set FieldNames = pFieldNames

End Property
'
' Result Handling
' ---------------
'
' ### Single Value
'
Public Function SingleResult() As Maybe

    ' Could use (x) => x(0,0)
    ' but less lambdas the better
    Dim result As Maybe
    If pRecords.IsSome Then
        Set result = Maybe.Some(pRecords.GetItem(0, 0))
    Else
        Set result = Maybe.None
    End If

    Set SingleResult = result
    
End Function
'
' ### Lists
'
Public Function GetColumn(ByVal colIndex As Long) As Variant()

    GetColumn = Application.index(pRecords, colIndex + 1, 0)

End Function
Public Function GetRow(ByVal rowIndex As Long) As Variant()

    GetRow = Application.index(pRecords, 0, rowIndex + 1)

End Function
'
' ### Dicts
'
''
' [ [$1], [$2] ] => { $1 -> $2 }
Public Function HDict() As Dict
    
    ' Check for only 2 fields
    Dim recs() As Variant
    recs = pRecords.GetItem
    
    If UBound(recs, 1) > 1 Then
        Err.Raise TOO_MANY_FIELDS, _
            description:=TypeName(Me) & ".HDict: Query returned more than 2 fields."
    End If
    
    Dim keys As List
    Set keys = List.Copy(GetColumn(0))
    
    Dim vals As List
    Set vals = List.Copy(GetColumn(1))
    
    Dim result As Dict
    Set result = Dict.FromLists(keys, vals)
    
    Set HDict = result
    
End Function
''
' () => { HEADER -> [ column ] }
Public Function VDict() As Dict
    
    Dim result As Dict
    Set result = Dict.Create
    
    Dim col As Integer
    For col = 1 To pFieldNames.Count
        result.Add pFieldNames(col), GetColumn(col - 1)
    Next
    
    Set VDict = result
    
End Function
''
' rowIndex => { HEADER -> value }
Public Function DictAt(ByVal rowIndex As Long) As Dict

    Set DictAt = Dict.FromLists(pFieldNames, List.Copy(GetRow(rowIndex)))

End Function
''
' () => [ { HEADER -> value } ]
Public Function AllDicts() As List
    
    ' TODO: Use a convert to instead of map with identity.
    Set AllDicts = ToBuildable(List)

End Function
'
' Interfaces
' ----------
'
' ### IShowable
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pFieldNames, pRecords)

End Function
Private Function IShowable_Show() As String
    
    IShowable_Show = Show
    
End Function
'
' ### IIterable
'
Private Function IIterable_Item(ByVal index As Long) As Variant

    Set IIterable_Item = DictAt(index)

End Function
Public Function LowerBound() As Long

    LowerBound = 0

End Function
Private Function IIterable_LowerBound() As Long

    IIterable_LowerBound = LowerBound

End Function
Public Function UpperBound() As Maybe

    Set UpperBound = pRecords.Map(Lambda.FromShort("UBound( _, 2 )"))

End Function
Private Function IIterable_UpperBound() As Long
    
    IIterable_UpperBound = UpperBound.GetOrElse(-1)

End Function
Public Function ToArray() As Variant()

    ToArray = defIterable.ToArray(Me)

End Function
Private Function IIterable_ToArray() As Variant()

    IIterable_ToArray = ToArray
    
End Function
Public Function ToCollection() As Collection

    Set ToCollection = defIterable.ToCollection(Me)

End Function
Private Function IIterable_ToCollection() As Collection

    Set IIterable_ToCollection = ToCollection

End Function
Public Function ToBuildable(ByVal seed As IBuildable) As IBuildable

    Set ToBuildable = defIterable.ToBuildable(seed, Me)

End Function
Private Function IIterable_ToBuildable(ByVal seed As IBuildable) As IBuildable

    Set IIterable_ToBuildable = ToBuildable(seed)

End Function
