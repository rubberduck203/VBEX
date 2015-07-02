VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lazy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' Lazy Class
' ============
'
'
' Copywrite (C) 2014 Philip Wales
' This file (Lazy.cls) is distributed under the MIT license
' Obtain a copy of the MIT license <http://opensource.org/licenses/MIT>
'
Implements IShowable
Implements IDelayed
'
' Private Members
' ---------------
'
Private pDelayed As IDelayed
Private pValue As Variant
'
' Constructors
' ------------
'
Private Sub Class_Initialize()

    Set pValue = Maybe.None

End Sub
Public Function Make(ByVal delayed As IDelayed) As Lazy

    Dim result As New Lazy
    Set result.DelayedOp = delayed
    Set Make = result

End Function
Friend Property Set DelayedOp(ByVal delayed As IDelayed)

    Set pDelayed = delayed

End Property
'
' Public Methods
' --------------
'
Public Function IsDelayed() As Boolean

    IsDelayed = pValue.IsNone

End Function
Public Function IsEvaluated() As Boolean

    IsEvaluated = pValue.IsSome

End Function
'
' IDelayed
' --------
'
Public Function Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    If IsDelayed Then
        Set pValue = Maybe.Some(pDelayed.Evaluate)
    End If

    Assign Evaluate, pValue.GetItem

End Function
Private Function IDelayed_Evaluate() As Variant

    Assign IDelayed_Evaluate, Evaluate

End Function
Public Function Map(ByVal op As IApplicable) As Lazy

    Dim result As IDelayed
    If IsDelayed Then
        Set result = pDelayed.Map(op)
    Else
        Set result = ByName.Create(op, pValue.GetItem)
    End If
    
    Set Map = Lazy.Make(result)

End Function
Private Function IDelayed_Map(ByVal op As IApplicable) As IDelayed

    Set IDelayed_Map = Map(op)

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Dim repr As String
    repr = defShow.Show(pValue.GetOrElse(pDelayed))

    Show = defShow.ParamShowableObject(Me, repr)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
