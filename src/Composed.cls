VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Composed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' Composed Class
' ==============
'
' Copywrite (C) 2014 Philip Wales
' This file (Composed.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Option Explicit
Implements IShowable
Implements IApplicable
'
' Private Members
' ---------------
'
Private pInnerOp As IApplicable
Private pOuterOp As IApplicable
'
' Constructors
' ------------
'
Public Function Make(ByVal outerOp As IApplicable, _
        ByVal innerOp As IApplicable) As Composed

    Dim result As New Composed
    Set result.inner = innerOp
    Set result.outer = outerOp
    Set Make = result

End Function
Friend Property Set inner(ByVal innerOp As IApplicable)

    Set pInnerOp = innerOp

End Property
Friend Property Set outer(ByVal outerOp As IApplicable)

    Set pOuterOp = outerOp

End Property
'
' IApplicable
' -----------
' 
' ### Meaningful
'
Public Function ApplyOnArray(ByRef args() As Variant) As Variant

    Assign ApplyOnArray, pOuterOp.Apply(pInnerOp.ApplyOnArray(args))

End Function
'
' ### Embedded
'
Private Function IApplicable_ApplyOnArray(ByRef args() As Variant) As Variant

    Assign IApplicable_ApplyOnArray, ApplyOnArray(args)
   
End Function
'! default member
Public Function Apply(ParamArray args() As Variant) As Variant
Attribute Apply.VB_UserMemId = 0

    Assign Apply, ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_Apply(ParamArray args() As Variant) As Variant

    Assign IApplicable_Apply, ApplyOnArray(CArray(args))
   
End Function
Public Function Compose(ByVal f As IApplicable) As IApplicable

    Set Compose = defApply.Compose(Me, f)

End Function
Private Function IApplicable_Compose(ByVal f As IApplicable) As IApplicable

    Set IApplicable_Compose = Compose(f)

End Function
Public Function AndThen(ByVal g As IApplicable) As IApplicable

    Set AndThen = defApply.AndThen(Me, g)

End Function
Private Function IApplicable_AndThen(ByVal g As IApplicable) As IApplicable

    Set IApplicable_AndThen = AndThen(g)

End Function
Public Function Partial(ParamArray args() As Variant) As IApplicable

    Set Partial = AsPartial(CArray(args))

End Function
Private Function IApplicable_Partial(ParamArray args() As Variant) As IApplicable

    Set IApplicable_Partial = AsPartial(Carray(args)

End Function
Public Function AsPartial(ParamArray args() As Variant) As IApplicable

    Set AsPartial = defApply.AsPartial(args)

End Function
Private Function IApplicable_AsPartial(ByRef args() As Variant) As IApplicable

    Set IApplicable_AsPartial = AsPartial(args)

End Function
Public Function Delay(ParamArray args() As Variant) As IDelayed

    Set Delay = AsDelay(CArray(args))

End Function
Private Function IApplicable_Delay(ParamArray args() As Variant) As IDelayed

    Set IApplicable_Delay = AsDelay(CArray(args))

End Function
Public Function AsDelay(ParamArray args() As Variant) As IDelayed

    Set AsDelay = defDelay.AsDelay(args)

End Function
Private Function IApplicable_AsDelay(ByRef args() As Variant) As IDelayed

    Set IApplicable_AsDelay = AsDelay(args)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pOuterOp, pInnerOp)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
