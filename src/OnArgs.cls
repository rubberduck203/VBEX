VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' OnArgs
' ======
'
' Copywrite (C) 2014 Philip Wales
' This file (OnArgs.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
' obj.method(_)
'
Option Explicit
Implements IApplicable
Implements IShowable
'
' Private Members
' ---------------
'
Private pObj As Object
Private pMethod As String
Private pCallType As VbCallType
'
' Constructors
' ------------
'
Public Function Make(ByVal method As String, ByVal callT As VbCallType, _
        ByVal obj As Object) As OnArgs

    Dim result As New OnArgs
    
    With result
        .MethodName = method
        .CallType = callT
        Set .Callee = obj
    End With
    
    Set Make = result

End Function
'
' ### Friend Methods
'
Friend Property Set Callee(ByVal obj As Object)
    
    Set pObj = obj
    
End Property
Friend Property Let MethodName(ByVal method As String)
    
    pMethod = method
    
End Property
Friend Property Let CallType(ByVal callT As String)
    
    pCallType = callT
    
End Property
'
' IApplicable
' -----------
'
Public Function ApplyOnArray(ByRef args() As Variant) As Variant

    Assign ApplyOnArray, defApply.CallByNameOnArray(pObj, pMethod, pCallType, args)

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
   
    Dim sObj As String
    sObj = defShow.Show(pObj)
    
    Dim repr As String
    repr = sObj & "." & pMethod
    
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function


