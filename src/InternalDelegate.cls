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
Implements Showable
Implements Applicable
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
' Applicable
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
Private Function Applicable_Apply(ParamArray args()) As Variant

    Assign Applicable_Apply, Applicable_ApplyOnArray(CArray(args))
    
End Function
Private Function Applicable_ApplyOnArray(ByRef args() As Variant) As Variant
    
    Assign Applicable_ApplyOnArray, ApplyOnArray(args)
    
End Function
Public Function Compose(ByVal f As Applicable) As Applicable

    Set Compose = defApply.Compose(Me, f)

End Function
Private Function Applicable_Compose(ByVal f As Applicable) As Applicable

    Set Applicable_Compose = Compose(f)

End Function
Public Function AndThen(ByVal g As Applicable) As Applicable

    Set AndThen = defApply.AndThen(Me, g)

End Function
Private Function Applicable_AndThen(ByVal g As Applicable) As Applicable

    Set Applicable_AndThen = AndThen(g)

End Function
Public Function Partial(ParamArray args() As Variant) As Applicable

    Set Partial = AsPartial(CArray(args))

End Function
Private Function Applicable_Partial(ParamArray args() As Variant) As Applicable

    Set Applicable_Partial = AsPartial(CArray(args))

End Function
Public Function AsPartial(ByRef args() As Variant) As Applicable

    Set AsPartial = defApply.AsPartial(Me, args)

End Function
Private Function Applicable_AsPartial(ByRef args() As Variant) As Applicable

    Set Applicable_AsPartial = AsPartial(args)

End Function
Public Function Delay(ParamArray args() As Variant) As delayed

    Set Delay = AsDelay(CArray(args))

End Function
Private Function Applicable_Delay(ParamArray args() As Variant) As delayed

    Set Applicable_Delay = AsDelay(CArray(args))

End Function
Public Function AsDelay(ByRef args() As Variant) As delayed

    Set AsDelay = defApply.AsDelay(Me, args)

End Function
Private Function Applicable_AsDelay(ByRef args() As Variant) As delayed

    Set Applicable_AsDelay = AsDelay(args)

End Function
'
' Showable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pFunctionName)

End Function
Private Function Showable_Show() As String

    Showable_Show = Show
    
End Function
