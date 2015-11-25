VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Partial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' Partial
' =======
'
' IApplicable with some args preset
'
Option Explicit
Implements IShowable
Implements IApplicable
'
' Constants
' ---------
'
'
' Private Members
' ---------------
'
Private pApplicable As IApplicable
Private pStoredArgs() As Variant
Private pArgLower As Long
Private pArgUpper As Long
'
' Constructors
' ------------
'
Public Function Make(ByVal app As IApplicable, ByRef storedArgs()) As Partial

    Dim result As New Partial
    Set result.Applicable = app
    result.StoreArguments storedArgs
    Set Make = result

End Function
Public Function Create(ByVal app As IApplicable, ParamArray storedArgs())

    Set Create = Make(app, CArray(storedArgs))

End Function
'
' Friend Methods
' --------------
'
Friend Property Set Applicable(ByVal app As IApplicable)
    
    Set pApplicable = app
    
End Property
Friend Sub StoreArguments(ByRef args())

    pStoredArgs = args
    pArgLower = LBound(pStoredArgs)
    pArgUpper = UBound(pStoredArgs)

End Sub
'
' IApplicable
' -----------
'
Public Function Apply(ParamArray args()) As Variant
Attribute Apply.VB_UserMemId = 0
On Error GoTo SimpleBubble

    Assign Apply, ApplyOnArray(CArray(args))

Exit Function
SimpleBubble:
    Exceptions.BubbleError Me, "Apply", Err

End Function
Public Function ApplyOnArray(ByRef args()) As Variant

    On Error GoTo ErrHandler

    Dim sentArgs() As Variant
    ReDim sentArgs(pArgLower To pArgUpper)
    
    Dim passedArgsLower As Long
    passedArgsLower = LBound(args)

    Dim passedArgsUpper As Long
    passedArgsUpper = UBound(args)

    ' Weave args into pStoredArgs as sentArgs
    Dim a As Long
    a = passedArgsLower

    Dim s As Long
    For s = pArgLower To pArgUpper

        If IsEmpty(pStoredArgs(s)) Then
            Assign sentArgs(s), args(a)
            a = a + 1
        Else
            Assign sentArgs(s), pStoredArgs(s)
        End If

    Next

    If a < passedArgsUpper Then
        On Error GoTo 0
            Exceptions.IllegalArgument Me, "ApplyOnArray", _
                "Called with too many arguments"
        On Error GoTo ErrHandler
    End If

    Assign ApplyOnArray, pApplicable.ApplyOnArray(sentArgs)

Exit Function
ErrHandler:
    Select Case Err.Number
    Case vbErrorNums.INDEX_ERROR
        Exceptions.IndexError Me, "ApplyOnArray", _
            "Not called with enough arguments."
    Case Else
        Exceptions.BubbleError Me, "ApplyOnArray", Err
    End Select

End Function
Private Function IApplicable_Apply(ParamArray args()) As Variant

    Assign IApplicable_Apply, ApplyOnArray(CArray(args))
    
End Function
Private Function IApplicable_ApplyOnArray(ByRef args()) As Variant
    
    Assign IApplicable_ApplyOnArray, ApplyOnArray(args)
    
End Function
Public Function Compose(ByVal f As IApplicable) As IApplicable

    Set Compose = defApply.Compose(Me, f)

End Function
Public Function AndThen(ByVal g As IApplicable) As IApplicable

    Set AndThen = defApply.AndThen(Me, g)

End Function
Private Function IApplicable_Compose(ByVal f As IApplicable) As IApplicable

    Set IApplicable_Compose = Compose(f)

End Function
Private Function IApplicable_AndThen(ByVal g As IApplicable) As IApplicable

    Set IApplicable_AndThen = AndThen(g)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pApplicable, pStoredArgs)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function

