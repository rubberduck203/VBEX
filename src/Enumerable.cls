VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Enum EnumerableError
    vbMethodNotSupportedError = 438
    EnumerableNotIntializedError = vbObjectError + 6500
    EnumerableInvalidArgumentError
End Enum

Private Const NotInitializedErrorMessage As String = "Collection Property Not Set"
Private Const InvalidArgumentErrorMessage As String = "Invalid Argument."

Private mCollection As Collection

Public Property Set Collection(obj As Variant)

    If TypeName(obj) = "Collection" Then
        Set mCollection = obj
    ElseIf TypeName(obj) = "Enumerable" Then
        Set mCollection = obj.Collection
    Else
        Set mCollection = New Collection
        Merge obj
    End If

End Property

Public Function Merge(collectionObject As Variant)
' Tries to convert any object passed in to a collection.
' This allows collection *like* objects such as Worksheets and Ranges.

On Error GoTo ErrHandler

    Dim element As Variant
    For Each element In collectionObject
        mCollection.Add element
    Next

Exit Function

ErrHandler:
    Const ObjectNotEnumerableMessage As String = "Object is not Enumerable."
    
    If Err.number = vbMethodNotSupportedError Then
        Err.Raise Err.number, TypeName(Me), InvalidArgumentErrorMessage & " " & ObjectNotEnumerableMessage, Err.HelpFile, Err.HelpContext
    Else
        ReRaiseError Err
    End If

End Function

Public Property Get Collection() As Collection
    Set Collection = mCollection
End Property

'Constructor
Public Function Create() As Enumerable
    Set Create = New Enumerable
End Function

' Collection Wrappers
Public Sub Add(Item, Optional key, Optional Before, Optional After)
    mCollection.Add Item, key, Before, After
End Sub

Public Sub Remove(index)
    mCollection.Remove index
End Sub

Public Function Count()
    Count = mCollection.Count
End Function

Public Function Item(Optional index)
Attribute Item.VB_UserMemId = 0
    Item = mCollection.Item(index)
End Function

' "Static" functions to be used with default instance of Enumerable

Public Function Range(ByVal startValue As Long, ByVal endValue As Long) As Enumerable 'Collection
    
    Set mCollection = New Collection
        
    Dim i As Long
    For i = startValue To endValue
        mCollection.Add i
    Next
    
    Set Range = New Enumerable
    Set Range.Collection = mCollection
End Function

Public Function Repeat(ByVal value, ByVal times As Long) As Enumerable
    Set mCollection = New Collection
    
    Dim i As Long
    For i = 1 To times
        mCollection.Add value
    Next
    
    Set Repeat = New Enumerable
    Set Repeat.Collection = mCollection
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
        Set NewEnum = mCollection.[_NewEnum]
End Property

' All of these functions work only on Collections whose items have a default value.
' If the items do not have a default value,
'   Runtime Error 438 "Object doesn't support this property or method" is raised.

' Instance Methods

Public Function Contains(itemToSearchFor As Variant, Optional ByVal compareByDefaultProperty = False) As Boolean
Attribute Contains.VB_Description = "Checks if an item exists in a Collection. Matches on the Default Property by Default. Runtime Error 438 'Object does not support method' may be raised when using 'compareByDefaultProperty'."
On Error GoTo ErrHandler

    Dim Item As Variant
    
    'compareByDefaultProperty is an unsafe option
    For Each Item In mCollection
        If IsObject(Item) And Not compareByDefaultProperty Then
            If Item Is itemToSearchFor Then
                Contains = True
                Exit Function
            End If
        Else
            If Item = itemToSearchFor Then
                Contains = True
                Exit Function
            End If
        End If
    Next Item
    
ExitFunction:
    Contains = False
    Exit Function

ErrHandler:
    HandleComparisonError Err
    Resume ExitFunction
End Function

Public Function First() As Variant
    First = mCollection(1)
End Function

Public Function Last() As Variant
    Last = mCollection(mCollection.Count)
End Function

Public Function Min() As Variant
On Error GoTo ErrHandler

    Dim Item As Variant
    Dim result As Variant
    
    For Each Item In mCollection
        If IsEmpty(result) Then
            result = Item
        ElseIf Item < result Then
            result = Item
        End If
    Next Item
    
    Min = result

ExitFunction:
    Exit Function
ErrHandler:
    HandleComparisonError Err
End Function

Public Function Max() As Variant
On Error GoTo ErrHandler
    
    Dim Item As Variant
    Dim result As Variant
    
    For Each Item In mCollection
        If IsEmpty(result) Then
            result = Item
        ElseIf Item > result Then
            result = Item
        End If
    Next Item
    
    Max = result
    
ExitFunction:
    Exit Function
ErrHandler:
    HandleComparisonError Err
End Function

Public Function Intersect(collection2 As Enumerable) As Enumerable 'Collection
    On Error GoTo ErrHandler
    
    If collection2 Is Nothing Then
        Err.Raise EnumerableInvalidArgumentError, TypeName(Me), InvalidArgumentErrorMessage
    End If
    
    Dim results As Enumerable
    Set results = New Enumerable
    Dim Item As Variant
    Dim innerItem As Variant
    
    For Each Item In mCollection
        For Each innerItem In collection2
            If Item = innerItem And Not IsEmpty(Item) Then
                If Not results.Contains(innerItem) Then 'curse the lack of shortcircuiting
                    results.Add innerItem
                    Exit For
                End If
            End If
        Next innerItem
    Next Item
    
    Set Intersect = results
    
ExitFunction:
    Exit Function
ErrHandler:
    HandleComparisonError Err
End Function

Public Function Distinct() As Enumerable
    Set Distinct = New Enumerable
    Set Distinct = Me.Intersect(Me)
End Function

Public Function Clone() As Enumerable
    Set Clone = New Enumerable
    Set Clone.Collection = CloneCollection
End Function

Public Function CloneCollection() As Collection
    Dim element As Variant
    Dim results As New Collection
    
    For Each element In mCollection
        results.Add Item
    Next
    
    Set CloneCollection = results
End Function

Public Function ToArray() As Variant
    Dim arr() As Variant
    ReDim arr(mCollection.Count - 1)
    Dim element As Variant
    
    Dim i As Long: i = 0
    For Each element In mCollection
        arr(i) = element
        i = i + 1
    Next
    
    ToArray = arr
End Function

Public Sub sort()
' implements a bubble sort
On Error GoTo ErrHandler
    
    Dim i As Long
    Dim j As Long
    Dim hasSwapped As Boolean
    Dim collectionCount As Long: collectionCount = mCollection.Count
    
    For i = collectionCount To 2 Step -1
        hasSwapped = False
        
        For j = 1 To i - 1
            If mCollection(j) > mCollection(j + 1) Then
                mCollection.Add mCollection(j), After:=j + 1
                mCollection.Remove j
                hasSwapped = True
            End If
        Next j

        If Not hasSwapped Then Exit For
    Next i
    
ExitFunction:
    Exit Sub
    
ErrHandler:
    HandleComparisonError Err
    
End Sub

Public Function IsSorted() As Boolean

    Dim Item As Variant
    Dim previous As Variant

    For Each Item In mCollection
        If Item < previous Then
            IsSorted = False
            Exit Function
        End If

        If IsObject(Item) Then
            Set previous = Item
        Else
            previous = Item
        End If
    Next Item

    IsSorted = True
End Function

Private Sub AssignUnknown(ByRef destination As Variant, ByRef source As Variant)
    If IsObject(source) Then
        Set destination = source
    Else
        destination = source
    End If
End Sub

Private Sub HandleComparisonError(error As ErrObject)

    Const ComparisonNotSupportedMessage As String = "An item in the collection does not have a default property; Cannot compare items without a default property."

    If error.number = vbMethodNotSupportedError Then
        error.Raise error.number, TypeName(Me), ComparisonNotSupportedMessage, error.HelpFile, error.HelpContext
    Else
        ReRaiseError error
    End If
    
End Sub

Private Sub ReRaiseError(error As ErrObject)
    error.Raise error.number, error.source, error.Description, error.HelpFile, error.HelpContext
End Sub

Private Sub Class_Initialize()
    Set mCollection = New Collection
End Sub
