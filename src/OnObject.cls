VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' OnObject
' ========
'
' Copywrite (C) 2014 Philip Wales
' This file (OnObject.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
' _.method(args)
'
Option Explicit
Implements IApplicable
Implements IShowable
'
' Private Members
' ---------------
'
Private pMethod As String
Private pCallType As VbCallType
Private pArgs() As Variant
'
' Constructors
' ------------
'
Public Function Create(ByVal method As String, ByVal callT As VbCallType _
        ParamArray args() As Variant) As OnObject

    Set Create = Make(method, callT, CArray(args))

End Function
Public Function Make(ByVal method As String, ByVal callT As VbCallType, _
        ByRef args() As Variant) As OnObject

    Dim result As New OnObject
    
    With result
        .MethodName = method
        .CallType = callT
        .Arguments = args
    End With
    
    Set Make = result

End Function
'
' ### Friend Methods
'
Friend Property Let MethodName(ByVal method As String)
    
    pMethod = method
    
End Property
Friend Property Let CallType(ByVal callT As String)
    
    pCallType = callT
    
End Property
Friend Property Let Arguments(ByRef args() As Variant)

    pArgs = args

End Property
'
' Errors
' ------
'
Private Sub InvalidArguementsError(ByRef params() As Variant)

    Dim failArray As String
    failArray = defShow.Show(params)
    
    Dim msg As String
    msg = "Invalid Arguements, Expecting one object " & failArray
    
    Exceptions.ValueError Me, "IApplicable_ApplyOnArray", msg
    
End Sub
'
' IApplicable
' -----------
'
Public Function ApplyOnArray(ByRef params() As Variant) As Variant

    Dim Callee As Object
    If UBound(params) = LBound(params) Then
        Set Callee = params(UBound(params))
    Else
        InvalidArguementsError params
    End If
    
    Assign ApplyOnArray, defApply.CallByNameOnArray(Callee, pMethod, pCallType, pArgs)

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
            
    Dim sArgs As String
    sArgs = defShow.Show(pArgs)
    
    Dim repr As String
    repr = "." & pMethod & "(" & sArgs & ")"
            
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
