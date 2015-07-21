VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IApplicable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Apply(ParamArray args() As Variant) As Variant
Attribute Apply.VB_UserMemId = 0

End Function
Public Function ApplyOnArray(ByRef args() As Variant) As Variant

End Function
Public Function Compose(ByVal f As IApplicable) As IApplicable

End Function
Public Function AndThen(ByVal g As IApplicable) As IApplicable

End Function

