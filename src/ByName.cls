VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' ByName Class
' ============
'
'
' Copywrite (C) 2014 Philip Wales
' This file (ByName.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
Implements IShowable
Implements IDelayed
'
' Private Members
' ---------------
'
Private pApplicable As IApplicable
Private pArgs() As Variant
'
' Constructors
' ------------
'
Public Function Create(ByVal op As IApplicable, ParamArray args()) As ByName

    Set Create = Make(op, CArray(args))

End Function
Public Function Make(ByVal op As IApplicable, ByRef args()) As ByName
    
    Dim result As New ByName
    Set result.Applicable = op
    result.Arguements = args
    Set Make = result

End Function
Friend Property Set Applicable(ByVal op As IApplicable)

    Set pApplicable = op

End Property
Friend Property Let Arguements(ByRef args())

    pArgs = args

End Property
'
' IDelayed
' --------
'
Public Property Get Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    Assign Evaluate, pApplicable.ApplyOnArray(pArgs)

End Property
Private Property Get IDelayed_Evaluate() As Variant

    Assign IDelayed_Evaluate, Evaluate

End Property
Public Function Map(ByVal op As IApplicable) As ByName

    Dim opResult As IApplicable
    Set opResult = op.Compose(pApplicable)
    
    Dim result As ByName
    Set result = Make(opResult, Tuple.Implode(pArgs))
    
    Set Map = result

End Function
Private Function IDelayed_Map(ByVal op As IApplicable) As IDelayed

    Set IDelayed_Map = Map(op)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pApplicable, pArgs)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
