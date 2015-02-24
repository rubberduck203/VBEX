Attribute VB_Name = "sort"
Option Explicit

' sort
' ====
'
' Contains in-place sorting algorithms for Arrays.
' ** ARRAYS ONLY **
'
' Helper Methods
' --------------
'
''
' Swap should work on an array or any two variables. It
' will not work on elements of sequence objects as the
' accessors of those return a value not a reference.
'
' x = "a": y = "b"
' Swap x, y ' x="b", y="a"
'
' a = Array("a", "b")
' Swap a(0), a(1) ' a = [b, a]
'
Private Sub Swap(ByRef x As Variant, ByRef y As Variant)
    
    Dim t As Variant
    cast.Assign t, x
    cast.Assign x, y
    cast.Assign y, t
    
End Sub
'
' Reversal
' --------
'
Public Sub Reverse(ByRef sequence() As Variant, _
        ByVal lower As Long, ByVal upper As Long)
    
    Do While lower < upper
        
        Swap sequence(lower), sequence(upper)
        
        lower = lower + 1
        upper = upper - 1
        
    Loop
    
End Sub
'
' Sorting
' -------
'
' ### Bubble Sort
'
Public Sub BubbleSort(ByRef sequence() As Variant, _
        ByVal lower As Long, ByVal upper As Long)

    Dim upperIt As Long
    For upperIt = upper To lower + 1 Step -1
        
        Dim hasSwapped As Boolean
        hasSwapped = False
        
        Dim bubble As Long
        For bubble = lower To upperIt - 1
            
            If sequence(bubble) > sequence(bubble + 1) Then
                
                Swap sequence(bubble), sequence(bubble + 1)
                hasSwapped = True
                
            End If
            
        Next bubble
        
        If Not hasSwapped Then Exit Sub
        
    Next upperIt
    
End Sub
'
' ### Quick Sort
'
Public Sub QuickSort(ByRef sequence() As Variant, ByVal lower As Long, ByVal upper As Long)
    
    ' length <= 1; already sorted
    If lower >= upper Then Exit Sub
    
    ' no special pivot selection used
    Swap sequence(seq.MiddleInt(lower, upper)), sequence(upper)
    
    ' pivot is at the end
    Dim pivot As Variant
    pivot = sequence(upper)
    
    Dim middle As Integer
    middle = Partition(sequence, lower, upper, pivot)
    
    ' don't swap if they are the same (pivot is single greatest)
    If middle <> upper Then Swap sequence(upper), sequence(middle)
    
    ' Omit the location of the pivot
    QuickSort sequence, lower, middle - 1
    
    ' it is exactly where it should be.
    QuickSort sequence, middle + 1, upper
    ' which is the magic of the quick sort
    
End Sub
Private Function Partition(ByRef sequence() As Variant, ByVal lower As Long, _
        ByVal upper As Long, ByVal pivot As Variant) As Long
        
    Do While lower < upper
        
        Do While sequence(lower) < pivot And lower < upper
            lower = lower + 1
        Loop
        
        ' right claims pivot as it is at the end
        Do While sequence(upper) >= pivot And lower < upper
            upper = upper - 1
        Loop
        
        ' don't swap if they are the same
        If lower <> upper Then Swap sequence(lower), sequence(upper)
        
    Loop
    Partition = lower
    
End Function
'
' ### Merge Sort?
'
'
' ### Insert Sort?
'
'
