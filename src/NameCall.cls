VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NameCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' NameCall Class
' ===========
'
'
' Copywrite (C) 2014 Philip Wales
' This file (NameCall.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
' These Should be two different classes but OH WELL!
'
' (obj).(method)(args)
'
'
Implements IApplicable
Implements IShowable
'
' Private Members
' ---------------
'
Private Enum Components
    obj
    method
    args
End Enum
Private pMissingComp As Components
Private pObj As Object
Private pMethod As String
Private pCallType As VbCallType
Private pArgs() As Variant
'
' Constructors
' ------------
'
Public Function OnArgs(ByVal method As String, ByVal callT As VbCallType, _
        ByVal obj As Object) As NameCall

    Dim result As New NameCall
    With result
        .MethodName = method
        .CallType = callT
        Set .callee = obj
        .MissingComp = Components.args
    End With
    Set OnArgs = result

End Function
Public Function OnObject(ByVal method As String, ByVal callT As VbCallType, _
        ParamArray args() As Variant) As NameCall

    Dim result As New NameCall
    With result
        .MethodName = method
        .CallType = callT
        .Arguments = cast.CArray(args)
        .MissingComp = Components.obj
    End With
    Set OnObject = result

End Function
'
' ### Friend Methods
'
Friend Property Set callee(ByVal obj As Object)
    
    Set pObj = obj
    
End Property
Friend Property Let MethodName(ByVal method As String)
    
    pMethod = method
    
End Property
Friend Property Let CallType(ByVal callT As String)
    
    pCallType = callT
    
End Property
Friend Property Let Arguments(ByRef args() As Variant)

    pArgs = args

End Property
Friend Property Let MissingComp(ByVal mComp As Components)

    pMissingComp = mComp

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
Public Function Apply(ParamArray args() As Variant) As Variant

    Assign Apply, IApplicable_ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_Apply(ParamArray args() As Variant) As Variant

    Assign IApplicable_Apply, IApplicable_ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_ApplyOnArray(ByRef params() As Variant) As Variant

    Dim args()
    Dim callee As Object
    
    Select Case pMissingComp
        Case Components.obj
            If UBound(params) = 0 Then
                Set callee = params(0)
                args = pArgs
            Else
                InvalidArguementsError params
            End If
        Case Components.args
            Set callee = pObj
            args = params
        Case Else
            NotImplementedError Me, "IApplicable_ApplyOnArray"
    End Select
    
    Dim result
    Select Case UBound(args) + 1
        Case 0
            Assign result, CallByName(callee, pMethod, pCallType)
        Case 1
            Assign result, CallByName(callee, pMethod, pCallType, args(0))
        Case 2
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1))
        Case 3
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2))
        Case 4
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3))
        Case 5
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4))
        Case 6
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4), args(5))
        Case 7
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
        Case 8
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 9
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 10
            Assign result, CallByName(callee, pMethod, pCallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case Else
            NotImplementedError Me, "IApplicable_ApplyOnArray"
    End Select

    Assign IApplicable_ApplyOnArray, result

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

    Dim repr As String
    Select Case pMissingComp
        Case obj
            
            Dim sArgs As String
            sArgs = defShow.Show(pArgs)
            repr = "." & pMethod & "(" & sArgs & ")"
            
        Case args
        
            Dim sObj As String
            sObj = defShow.Show(pObj)
            repr = sObj & "." & pMethod
            
    End Select
    
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
