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
Implements IMonadic
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
'
' IMonadic
' --------
'
Public Function Bind(ByVal op As IApplicable) As Lazy

    Dim result As Lazy
    Set result = op.Apply(pDelayed)
    
    Set Bind = result

End Function
Public Function Map(ByVal op As IApplicable) As Lazy

    Dim x As Lazy
    Set x = Me
    
    Set Map = Lazy.Make(ByName.Make(op, Tuple.Pack(x)))

End Function
Public Function Unit(ByVal x) As Lazy

    Set Unit = Make(x)

End Function
Private Function IMonadic_Bind(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Bind = Bind(op)

End Function
Private Function IMonadic_Map(ByVal op As IApplicable) As IMonadic

    Set IMonadic_Map = Map(op)

End Function
Private Function IMonadic_Unit(ByVal x) As IMonadic

    Set IMonadic_Unit = Unit(x)

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
