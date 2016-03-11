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
' This file (Lazy.cls) is distributed under the GPL-3.0 license
' Obtain a copy of the GPL-3.0 license <http://opensource.org/licenses/GPL-3.0>
'
Implements Showable
Implements delayed
'
' Private Members
' ---------------
'
Private pDelayed As delayed
Private pValue As Variant
'
' Constructors
' ------------
'
Public Function Create(ByVal op As Applicable, ParamArray args()) As Lazy

    Set Create = Make(op, CArray(args))

End Function
Public Function Make(ByVal op As Applicable, ByRef args()) As Lazy
    
    Dim delayed As ByName
    Set delayed = ByName.Make(op, args)

    Set Make = FromDelayed(delayed)

End Function
Public Function FromDelayed(ByVal delayed As delayed) As Lazy

    Dim result As New Lazy
    Set result.DelayedOp = delayed
    Set FromDelayed = result
    
End Function
Friend Property Set DelayedOp(ByVal delayed As delayed)

    Set pDelayed = delayed

End Property
'
' Public Methods
' --------------
'
Public Function IsDelayed() As Boolean

    IsDelayed = IsEmpty(pValue)

End Function
Public Function IsEvaluated() As Boolean

    IsEvaluated = (Not IsDelayed)

End Function
'
' Delayed
' --------
'
Public Property Get Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0

    If IsDelayed Then
        Assign pValue, pDelayed.Evaluate
    End If

    Assign Evaluate, pValue

End Property
Private Property Get Delayed_Evaluate() As Variant

    Assign Delayed_Evaluate, Evaluate

End Property
Public Function Map(ByVal op As Applicable) As Lazy

    Dim result As delayed
    If IsDelayed Then
        ' `Me` in a parameter array evauluates the default property...
        Dim args(0 To 0) As Variant
        Assign args(0), Me
        Set result = ByName.Make(op, args) ' Evaluating result does not evaluate `Me` when using a CallByName Applicable.
        Debug.Assert IsDelayed
    Else
        Set result = ByName.Create(op, pValue)
    End If
    
    Set Map = Lazy.FromDelayed(result)

End Function
Private Function Delayed_Map(ByVal op As Applicable) As delayed

    Set Delayed_Map = Map(op)

End Function
'
' Showable
' ---------
'
Public Function Show() As String

    Dim repr As String
    If IsDelayed Then
        repr = defShow.Show(pDelayed)
    Else
        repr = defShow.Show(pValue)
    End If
    Show = defShow.ParamShowableObject(Me, repr)

End Function
Private Function Showable_Show() As String

    Showable_Show = Show

End Function
