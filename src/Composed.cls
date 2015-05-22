VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Composed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Composed Class
' ==============
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Composed.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
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
    Set result.Inner = innerOp
    Set result.Outer = outerOp
    Set Make = result

End Function
Friend Property Set Inner(ByVal innerOp As IApplicable)

    Set pInnerOp = innerOp

End Property
Friend Property Set Outer(ByVal outerOp As IApplicable)

    Set pOuterOp = outerOp

End Property
'
' IApplicable
' -----------
'
Public Function Apply(ParamArray args() As Variant) As Variant

    Assign Apply, IApplicable_ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_Apply(ParamArray args() As Variant) As Variant

    Assign IApplicable_Apply, IApplicable_ApplyOnArray(cast.CArray(args))
   
End Function
Private Function IApplicable_ApplyOnArray(ByRef args() As Variant) As Variant

    Assign IApplicable_ApplyOnArray, pOuterOp.Apply(pInnerOp.ApplyOnArray(args))
   
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

    Show = defShow.ParamShowableObject(Me, pOuterOp, pInnerOp)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
