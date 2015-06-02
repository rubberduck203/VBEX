VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' OnArgs
' ======
'
'
' Copywrite (C) 2014 Philip Wales
' This file (OnArgs.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
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
Public Function Apply(ParamArray args() As Variant) As Variant

    Assign Apply, IApplicable_ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_Apply(ParamArray args() As Variant) As Variant

    Assign IApplicable_Apply, IApplicable_ApplyOnArray(cast.CArray(args))

End Function
Private Function IApplicable_ApplyOnArray(ByRef args() As Variant) As Variant

    Assign IApplicable_ApplyOnArray, defApply.CallByNameWithArray(pObj, pMethod, pCallType, args)

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
   
    Dim sObj As String
    sObj = defShow.Show(pObj)
    
    Dim repr As String
    repr = sObj & "." & pMethod
    
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function


