Attribute VB_Name = "sort"
Option Explicit

' sort
' ====
'
' Sort
' ----
'
' ### Bubble Sort
'
Public Sub BubbleSort(ByRef sequence As Variant, _
        ByVal lower As Long, ByVal upper As Long)

    Dim upperIt As Long
    For upperIt = upper To lower + 1 Step -1
        
        Dim hasSwapped As Boolean
        hasSwapped = False
        
        Dim bubble As Long
        For bubble = lower To upperIt - 1
            
            If sequence(bubble) > sequence(bubble + 1) Then
                
                seq.SwapIndexes sequence, bubble, bubble + 1
                hasSwapped = True
                
            End If
            
        Next bubble
        
        If Not hasSwapped Then Exit Sub
        
    Next upperIt
    
End Sub
'
' ### Quick Sort
'
Public Sub QuickSort(ByRef sequence As Variant, ByVal lower As Long, ByVal upper As Long)
    
    ' length <= 1; already sorted
    If lower >= upper Then Exit Sub
    
    ' no special pivot selection used
    SwapIndexes sequence, seq.MiddleInt(lower, upper), upper
    
    ' pivot is at the end
    Dim pivot As Variant
    pivot = sequence(upper)
    
    Dim middle As Integer
    middle = Partition(sequence, lower, upper, pivot)
    
    ' don't swap if they are the same (pivot is single greatest)
    If middle <> upper Then seq.SwapIndexes sequence, upper, middle
    
    ' Omit the location of the pivot
    QuickSort sequence, lower, middle - 1
    
    ' it is exactly where it should be.
    QuickSort sequence, middle + 1, upper
    ' which is the magic of the quick sort
    
End Sub
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
        If lower <> upper Then seq.SwapIndexes sequence, lower, upper
        
    Wend
    Partition = lower
    
End Function
'
' ### Merge Sort?
'
'
' ### Insert Sort?
'
'
