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
Implements Applicable
Implements Showable
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
Public Function Create(ByVal method As String, ByVal callT As VbCallType, _
        ParamArray args() As Variant) As OnObject

    Set Create = Make(method, callT, CArray(args))

End Function
Public Function Make(ByVal method As String, ByVal callT As VbCallType, _
        ByRef args() As Variant) As OnObject

    Dim result As New OnObject
    
    With result
        .MethodName = method
        .CallType = callT
        .LetArguments args
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
Friend Sub LetArguments(ByRef args() As Variant)

    pArgs = args

End Sub
'
' Errors
' ------
'
Private Sub InvalidArguementsError(ByRef params() As Variant)

    Dim failArray As String
    failArray = defShow.Show(params)
    
    Dim msg As String
    msg = "Invalid Arguements, Expecting one object " & failArray
    
    Exceptions.ValueError Me, "Applicable_ApplyOnArray", msg
    
End Sub
'
' Applicable
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
Private Function Applicable_ApplyOnArray(ByRef args() As Variant) As Variant

    Assign Applicable_ApplyOnArray, ApplyOnArray(args)
   
End Function
'! default member
Public Function Apply(ParamArray args() As Variant) As Variant
Attribute Apply.VB_UserMemId = 0

    Assign Apply, ApplyOnArray(cast.CArray(args))

End Function
Private Function Applicable_Apply(ParamArray args() As Variant) As Variant

    Assign Applicable_Apply, ApplyOnArray(CArray(args))
   
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
            
    Dim sArgs As String
    sArgs = defShow.Show(pArgs)
    
    Dim repr As String
    repr = "." & pMethod & "(" & sArgs & ")"
            
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function Showable_Show() As String

    Showable_Show = Show

End Function
