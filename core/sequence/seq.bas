Attribute VB_Name = "seq"
Option Explicit

' seq
' ===
'
' Sequence Helper functions and routines.
'
' Compatibility
' -------------
' Ensure compatibility with Arrays and other sequence types
'
' ### Bounds
'
''
' UpperBound: Sequence Objects are assumed to have a `.Count`
' property.
Public Function UpperBound(ByRef sequence As Variant) As Long
    
    On Error GoTo NotArray
    UpperBound = UBound(sequence)
    
CleanExit:
    Exit Function
    
NotArray:
    UpperBound = sequence.Count
    Resume CleanExit
    
End Function
''
' LowerBound: Sequence Objects are assumed to have 1 base offset!
Public Function LowerBound(ByRef sequence As Variant) As Long
    
    On Error GoTo NotArray
    LowerBound = LBound(sequence)
    
CleanExit:
    Exit Function
    
NotArray:
    LowerBound = 1
    Resume CleanExit
    
End Function

Public Function Length(ByVal sequence As Variant) As Long
    Length = UpperBound(sequence) - LowerBound(sequence) + 1
End Function
'
' Just Great Helper Functions
' ---------------------------
'
' Probably better to put it in another module but I think
' this will be the root of everything.
'
Public Sub Assign(ByRef x As Variant, ByVal y As Variant)
    
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
    
End Sub
''
' Can easily be extracted to use Floor(Average(a,b)) but I would like one 
' way dependence.
Private Function MiddleInt(ByVal a As Variant, ByVal b As Variant) As Variant
    MiddleInt = (a + b) \ 2
End Function

' Swapping
' --------

''
' Swap should work on an array or any two variables. It
' will not work on elements of sequence objects as the
' accessors of those return a value not a reference. For
' those use `SwapIndexes`.
'
' x = "a": y = "b"
' Swap x, y ' x="b", y="a"
'
' a = Array("a", "b")
' Swap a(0), a(1) ' a = [b, a]
'
Public Sub Swap(ByRef x As Variant, ByRef y As Variant)
    
    Dim t As Variant
    
    Assign t, x
    Assign x, y
    Assign y, t
    
End Sub
''
' `SwapIndexes` is to be used on sequence objects instead of `Swap`.  It uses the 
' default property of the object to access the elements.  If the default property
' is read-only an error is raised.  This cannot be used with a collection.
Public Sub SwapIndexes(ByRef sequence As Variant, ByVal a As Long, ByVal b As Long)
    ' We cannot implement `Assign` in case sequence is an object as `Collection.Item`
    ' returns a value and not a reference. Therefore we must reuse the pattern
    
    On Error GoTo IsCollection
    
    Dim t As Variant
    Assign t, sequence(a)
        
    If IsObject(sequence(b)) Then
        Set sequence(a) = sequence(b)
    Else
        sequence(a) = sequence(b)
    End If
    
    If IsObject(t) Then
        Set sequence(b) = t
    Else
        sequence(b) = t
    End If
    
CleanExit:
    Exit Sub
IsCollection:
    
    Err.Raise 13, "seq.SwapIndexes", "Sequence's default property is not read-write."
    
End Sub
'
' Enumeration
' -----------
'
' forgive the single letters I can't find an appropriate name for b that isn't taken.
Public Function Enumeration(ByVal a As Long, ByVal b As Long) As List
    
    Set Enumeration = New List
    
    Dim s As Integer
    s = IIF((a < b), 1, -1)
    
    Dim n As Long
    For n = a To b Step s
        Enumeration.Append n
    Next n
    
End Function

Public Function Enumerated(ByVal sequence As Variant) As List
    Set Enumerated = Enumeration(LowerBound(sequence), UpperBound(sequence))
End Function

Public Sub Fill(ByRef sequence As Variant, ByVal filler As Variant)
    
    Dim i As Long
    For i = LowerBound(sequence) To UpperBound(sequence)
        sequence(i) = filler
    Next i
    
End Sub
'
' Comparison
' ----------
'
Public Function Compare(ByRef seqA As Variant, ByRef seqB As Variant) As Boolean
    
    Compare = False

    Dim lowA As Long
    lowA = LowerBound(seqA)
    
    Dim upA As Long
    upA = UpperBound(seqA)

    If Not ((upA - lowA) = (UpperBound(seqB) - LowerBound(seqB))) Then
        Exit Function
    End If
    
    Dim offset As Long
    offset = LowerBound(seqB) - lowA
    
    Dim i As Long
    For i = lowA To upA
        If Not (seqA(i) = seqB(i + offset)) Then Exit Function
    Next i
    
    Compare = True
    
End Function
'
' Reversal
' --------
'
' Reverse should accept an array or collection
Public Sub Reverse(ByRef sequence As Variant)
    
    Dim lower_it As Long
    lower_it = LowerBound(sequence)
    
    Dim upper_it As Long
    upper_it = UpperBound(sequence)
    
    While lower_it < upper_it
        
        SwapIndexes sequence, lower_it, upper_it
        
        lower_it = lower_it + 1
        upper_it = upper_it - 1
        
    Wend
    
End Sub
''
' Reversed will return the type that was passed
' Strong typing seems silly now...
Public Function Reversed(ByVal sequence As Variant) As Variant
    
    Assign Reversed, sequence
    Reverse Reversed
    
End Function
'
' [head,t,a,i,l]
' --------------
'
Public Function Head(ByVal sequence As Variant) As Variant
    
    Assign Head, sequence(LowerBound(sequence))
    
End Function
'
' [i,n,i,t,last]
' --------------
'
Public Function Last(ByRef sequence As Variant) As Variant
    
    Assign Last, sequence(UpperBound(sequence))
    
End Function
'
' Search
' ------
'
' ### Maximums
'
''
' MaxIndex: Returns the index of `sequence` that has the maximum value
Public Function MaxIndex(ByRef sequence As Variant) As Long
    
    MaxIndex = LowerBound(sequence)
    
    Dim i As Long
    For i = LowerBound(sequence) To UpperBound(sequence)
        
        If sequence(MaxIndex) < sequence(i) Then MaxIndex = i
    
    Next i
    
End Function
''
' MaxValue: Returns the value of `sequence` that is the Maximum
' Uses `MaxIndex`
Public Function MaxValue(ByRef sequence As Variant) As Variant
    
    Assign MaxValue, sequence(MaxIndex(sequence))
    
End Function
''
' LinearSearch:
Public Function LinearSearch(ByVal value As Variant, sequence As Variant) As Long
    
    Dim i As Long
    For i = LowerBound(sequence) To UpperBound(sequence)
        
        If sequence(i) = value Then
            
            LinearSearch = i
            Exit Function
        
        End If
        
    Next i
    
    LinearSearch = -1
    
End Function
''
' Binary Search: Sequence must be sorted.  Has the option of returning where the
' value should be instead of not found.
Public Function BinarySearch(ByVal value As Variant, ByRef sorted_seq As Variant, _
                             Optional ByVal nearest As Boolean = False) As Long
    
    Dim upper As Long
    upper = UpperBound(sorted_seq)
    
    Dim lower As Long
    lower = LowerBound(sorted_seq)
    
    While lower < upper
        
        Dim middle As Long
        middle = MiddleInt(lower, upper)
        
        If sorted_seq(middle) >= value Then
            upper = middle
        Else
            lower = middle + 1
        End If
        
    Wend
    
    BinarySearch = IIF(sorted_seq(upper) = value Or nearest, upper, -1)
    
End Function
'
' Sort
' ----
'
' ### Bubble Sort
'
Public Sub BubbleSort(ByRef sequence As Variant)
    
    Dim lower As Long
    lower = LowerBound(sequence)
    
    Dim upper As Long
    For upper = UpperBound(sequence) To lower + 1 Step -1
        
        Dim hasSwapped As Boolean
        hasSwapped = False
        
        Dim bubble As Long
        For bubble = lower To upper - 1
            
            If sequence(bubble) > sequence(bubble + 1) Then
                
                 SwapIndexes sequence, bubble, bubble + 1
                 hasSwapped = True
                 
            End If
            
        Next bubble
        
        If Not hasSwapped Then Exit Sub
        
    Next upper
    
End Sub
Public Function BubbleSorted(ByVal sequence As Variant) As Variant
    
    Assign BubbleSorted, sequence
    BubbleSort BubbleSorted
    
End Function
'
' ### Insert Sort
'
Public Sub InsertSort(ByRef sequence As Variant)
    
    Dim lower As Long
    lower = LowerBound(sequence)
    
    Dim i As Long
    For i = lower + 1 To UpperBound(sequence)
        
        Dim value As Variant
        value = sequence(i)
        
        Dim j As Long
        j = i - 1
        While j >= lower And sequence(j) > value
        
            sequence(j + 1) = sequence(j)
            j = j - 1
        
        Wend
        
        sequence(j + 1) = value
        
    Next i
    
End Sub
Public Function InsertSorted(ByVal sequence As Variant) As Variant
    
    Assign InsertSorted, sequence
    InsertSort InsertSorted
    
End Function
'
' ### Quick Sort
'
Public Sub QuickSort(ByRef sequence As Variant, ByVal lower As Long, ByVal upper As Long)
    
    ' length <= 1; already sorted
    If lower >= upper Then Exit Sub
    
    ' no special pivot selection used
    SwapIndexes sequence, MiddleInt(lower, upper), upper
    
    ' pivot is at the end
    Dim pivot As Variant
    pivot = sequence(upper)
    
    Dim middle As Integer
    middle = Partition(sequence, lower, upper, pivot)
    
    ' don't swap if they are the same (pivot is single greatest)
    If middle <> upper Then SwapIndexes sequence, upper, middle
    
    ' Omit the location of the pivot
    QuickSort sequence, lower, middle - 1
    
    ' it is exactly where it should be.
    QuickSort sequence, middle + 1, upper
    ' which is the magic of the quick sort
    
End Sub
Public Function QuickSorted(ByVal sequence As Variant) As Variant
    
    Assign QuickSorted, sequence
    QuickSort QuickSorted, LowerBound(QuickSorted), UpperBound(QuickSorted)
    
End Function
Private Function Partition(ByRef sequence As Variant, ByVal lower As Long, _
                           ByVal upper As Long, ByVal pivot As Variant) As Long
    
    While lower < upper
        
        While sequence(lower) < pivot And lower < upper
            lower = lower + 1
        Wend
        
        ' right claims pivot as it is at the end
        While sequence(upper) >= pivot And lower < upper
            upper = upper - 1
        Wend
        
        ' don't swap if they are the same
        If lower <> upper Then SwapIndexes sequence, lower, upper
        
    Wend
    
    Debug.Assert lower = upper
    Partition = lower
    
End Function
'
' High Order Functions
' --------------------
'
' `Application.Run` has all arguements passed by value.  Ergo it should never be
' run without returning a value
'
Public Function Map(ByVal delegate As String, ByRef sequence As Variant) As List
    
    Dim mappedList As New List
    
    Dim i As Long
    For i = LowerBound(sequence) To LowerBound(sequence)
        mappedList.Append Application.Run(delegate, sequence(i))
    Next i
    
    Set Map = mappedList
    
End Function
Function DifficultMap(ByVal sequence As Variant, ByVal delegate As String) As List

    Dim answers As Object
    Set answers = CreateObject("Scripting.Dictionary")
    
    Dim mappedList As New List
    
    Dim i As Long
    For i = seq.LowerBound(sequence) To seq.UpperBound(sequence)
    
        If Not answers.Exists(sequence(i)) Then
        
            answers.Add key:=sequence(i), _
                        value:=Application.Run(delegate, sequence(i))
            mappedList.Append answers(sequence(i))
            
        Else
            mappedList.Append answers(sequence(i))
        End If
        
    Next i
    
    Set DifficultMap = mappedList
    
End Function
Public Function Fold(ByVal delegate As String, ByVal sequence As Variant, _
                     ByVal initial_value As Variant) As Variant
    
    Fold = initial_value
    
    Dim el As Variant
    For Each el In sequence
        Fold = Application.Run(delegate, Fold, el)
    Next el
    
End Function
Public Function Reduce(ByVal delegate As String, ByVal sequence As Variant) As Variant
    
    Reduce = Fold(delegate, Tail(sequence), Head(sequence))
    
End Function
Public Function Compose(ByVal delegates As Variant, ByVal initial_value As Variant) As Variant
    
    Compose = initial_value
    
    Dim delegate As String
    For Each delegate In delegates
        Compose = Application.Run(delegate, Compose)
    Next delegate
    
End Function
'
' Other Functions
' ---------------
'
Public Function Any_(ByVal sequence As Variant) As Boolean
    
    Any_ = True

    Dim element As Variant
    For Each element In sequence
        If element Then Exit Function
    Next element
    
    Any_ = False

End Function
Public Function All(ByVal sequence As Variant) As Boolean
    
    All = False

    Dim element As Variant
    For Each element In sequence
        If Not element Then Exit Function
    Next element
    
    All = True
    
End Function
Public Function Same(ByVal sequence As Variant) As Boolean
    
    Dim i As Variant
    For i = LowerBound(sequence) + 1 To UpperBound(sequence)
        Same = (sequence(i) = sequence(i - 1))
        If Not Same Then Exit Function
    Next i
    
End Function
Public Function ToArray(ByVal sequence As Variant) As Variant
    
    On Error GoTo EmptyCollection
    
    ' zero offset is enforced.
    Dim arr() As Variant
    ReDim arr(Length(sequence) - 1)
    
    Dim i As Long
    Dim element As Variant
    For Each element In sequence
        
        Assign arr(i), element
        i = i + 1
        
    Next element

CleanExit:
    ToArray = arr
    
    Exit Function
    
EmptyCollection:
    Err.Clear
    Debug.Assert sequence.Count = 0
    arr = Array()
    Resume CleanExit

End Function
