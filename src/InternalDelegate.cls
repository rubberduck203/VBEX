VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InternalDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' InternalDelegate
' ================
'
' Copywrite (C) 2014 Philip Wales
' This file (InternalDelegate.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
'
Option Explicit
Implements IShowable
Implements IApplicable
'
' Constants
' ---------
'
Private Const CANNOT_FIND_MACRO As Integer = 1004
'
' Private Members
' ---------------
'
Private pFunctionName As String
'
' Constructors
' ------------
'
Public Function Make(ByVal funcName As String) As InternalDelegate

    Dim result As New InternalDelegate
    result.FunctionName = funcName
    Set Make = result

End Function
'
' Friend Methods
' --------------
'
Friend Property Let FunctionName(ByVal funcName As String)

    pFunctionName = funcName

End Property
'
' IApplicable
' -----------
'
' ### Meaningful
'
'! default member
Public Function Apply(ParamArray args()) As Variant
Attribute Apply.VB_UserMemId = 0
On Error GoTo ErrHandler

    Assign Apply, ApplyOnArray(CArray(args))

Exit Function
ErrHandler:
    Select Case Err.Number

        Case CANNOT_FIND_MACRO

            Dim msg As String
            msg = pFunctionName & ": is not a global Function or Sub in VBEX"
            Exceptions.ValueError Me, "Apply", msg

        Case Else
            Exceptions.BubbleError Me, "Apply", Err

    End Select
End Function
'
' ### Embedded
'
Public Function ApplyOnArray(ByRef args() As Variant) As Variant

    Assign ApplyOnArray, defApply.ApplicationRunOnArray(pFunctionName, args)

End Function
Private Function IApplicable_Apply(ParamArray args()) As Variant

    Assign IApplicable_Apply, IApplicable_ApplyOnArray(CArray(args))
    
End Function
Private Function IApplicable_ApplyOnArray(ByRef args() As Variant) As Variant
    
    Assign IApplicable_ApplyOnArray, ApplyOnArray(args)
    
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

    Set IApplicable_Partial = AsPartial(CArray(args))

End Function
Public Function AsPartial(ByRef args() As Variant) As IApplicable

    Set AsPartial = defApply.AsPartial(Me, args)

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
Public Function AsDelay(ByRef args() As Variant) As IDelayed

    Set AsDelay = defApply.AsDelay(Me, args)

End Function
Private Function IApplicable_AsDelay(ByRef args() As Variant) As IDelayed

    Set IApplicable_AsDelay = AsDelay(args)

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
