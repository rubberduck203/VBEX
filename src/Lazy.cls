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
Public Function Make(ByVal delayed As IDelayed) As Lazy

    Dim result As New Lazy
    Set result.DelayedOp = delayed
    Set Make = result

End Function
Friend Property Set DelayedOp(ByVal delayed As IDelayed)

    Set pDelayed = delayed

End Property
'
' Public Functions
' ----------------
'
Public Function IsDelayed() As Boolean

    IsDelayed = IsEmpty(pValue)

End Function
Public Function IsEvaluated() As Boolean

    IsEvaluated = Not IsDelayed

End Function
'
' IDelayed
' --------
'
Public Function Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    If IsDelayed Then
        Assign pValue, pDelayed.Evaluate
    End If

    Assign Evaluate, pValue

End Function
Private Function IDelayed_Evaluate() As Variant

    Assign IDelayed_Evaluate, Evaluate

End Function
'
' IShowable
' ---------
'
Public Function Show() As String

    Dim repr As String
    If IsEvaluated Then
        repr = defShow.Show(pValue)
    Else
        repr = defShow.Show(pDelayed)
    End If

    Show = defShow.ParamShowableObject(Me, repr)

End Function
Private Function IShowable_Show() As String

    IShowable_Show = Show

End Function
