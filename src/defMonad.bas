Attribute VB_Name = "defMonad"
Option Explicit

Public Function ApplyUnit(ByVal m As IMonadic) As NameCall

    Set ApplyUnit = NameCall.OnArgs("Unit", VbMethod, m)

End Function
