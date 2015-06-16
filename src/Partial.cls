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
Private pArgs As Tuple
'
' Constructors
' ------------
'
Public Function Make(ByVal app As IApplicable, ByVal args As Tuple) As Partial

End Function
Public Function Prepend(ByVal app As IApplicable, ParamArray frontArgs()) As Partial

End Function
Public Function Append(ByVal app As IApplicable, ParamArray backArgs()) As Partial
    
End Function
'
' Friend Methods
' --------------
'
Friend Property Set Applicable(ByVal app As IApplicable)
    
    Set pApplicable = app
    
End Property
Friend Property Set Arguments(ByVal args As Tuple)

    Set pArgs = args

End Property
'
' IApplicable
' -----------
'
Public Function Apply(ParamArray args()) As Variant

    Assign Apply, IApplicable_ApplyOnArray(CArray(args))

End Function
Private Function IApplicable_Apply(ParamArray args()) As Variant

    Assign IApplicable_Apply, IApplicable_ApplyOnArray(CArray(args))
    
End Function
Private Function IApplicable_ApplyOnArray(ByRef args() As Variant) As Variant
    
    Dim sentArgs() As Variant
    ' sentArgs = ???
    Assign IApplicable_ApplyOnArray, pApplicable.ApplyOnArray(sentArgs)
    
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

    Show = defShow.ParamShowableObject(Me, pFunctionName)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show
    
End Function

