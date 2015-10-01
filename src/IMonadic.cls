VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMonadic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Bind(ByVal op As IApplicable) As IMonadic

End Function
Public Function Map(ByVal op As IApplicable) As IMonadic

End Function
Public Function Unit(ByVal x) As IMonadic
Attribute Unit.VB_UserMemId = 0

End Function
