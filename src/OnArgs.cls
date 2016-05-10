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
Implements Applicable
Implements Showable
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
    result.Inject obj, method, callT
    Set Make = result

End Function
'
' ### Friend Methods
'
Friend Sub Inject(ByVal obj As Object, ByVal method As String, ByVal callT As String)
    
    Set pObj = obj
    pMethod = method
    pCallType = callT
    
End Sub
'
' Applicable
' -----------
'
Public Function ApplyOnArray(ByRef args() As Variant) As Variant

    Assign ApplyOnArray, defApply.CallByNameOnArray(pObj, pMethod, pCallType, args)

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
Public Function delay(ParamArray args() As Variant) As Delayed

    Set delay = AsDelay(CArray(args))

End Function
Private Function Applicable_Delay(ParamArray args() As Variant) As Delayed

    Set Applicable_Delay = AsDelay(CArray(args))

End Function
Public Function AsDelay(ByRef args() As Variant) As Delayed

    Set AsDelay = defApply.AsDelay(Me, args)

End Function
Private Function Applicable_AsDelay(ByRef args() As Variant) As Delayed

    Set Applicable_AsDelay = AsDelay(args)

End Function
'
' Showable
' ---------
'
Public Function Show() As String
   
    Dim sObj As String
    sObj = defShow.Show(pObj)
    
    Dim repr As String
    repr = sObj & "." & pMethod
    
    Show = defShow.ParamShowableObject(Me, repr)
    
End Function
Private Function Showable_Show() As String

    Showable_Show = Show

End Function


