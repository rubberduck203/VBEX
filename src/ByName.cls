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
' This file (ByName.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Implements Showable
Implements delayed
'
' Private Members
' ---------------
'
Private pApplicable As Applicable
Private pArgs() As Variant
'
' Constructors
' ------------
'
Public Function Create(ByVal op As Applicable, ParamArray args()) As ByName

    Set Create = Make(op, CArray(args))

End Function
Public Function Make(ByVal op As Applicable, ByRef args()) As ByName
    
    Dim result As New ByName
    result.SetOp op
    result.SetArgs args
    Set Make = result

End Function
Friend Sub SetOp(ByVal op As Applicable)

    Set pApplicable = op

End Sub
Friend Sub SetArgs(ByRef args())

    pArgs = args

End Sub
'
' Delayed
' --------
'
Public Property Get Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    Assign Evaluate, pApplicable.ApplyOnArray(pArgs)

End Property
Private Property Get Delayed_Evaluate() As Variant

    Assign Delayed_Evaluate, Evaluate

End Property
Public Function Map(ByVal op As Applicable) As ByName

    Dim opResult As Applicable
    Set opResult = op.Compose(pApplicable)
    
    Dim result As ByName
    Set result = Make(opResult, pArgs)
    
    Set Map = result

End Function
Private Function Delayed_Map(ByVal op As Applicable) As delayed

    Set Delayed_Map = Map(op)

End Function
'
' Showable
' ---------
'
Public Function Show() As String

    Show = defShow.ParamShowableObject(Me, pApplicable, pArgs)

End Function
Private Function Showable_Show() As String

    Showable_Show = Show

End Function
