Attribute VB_Name = "TestVbeCodeModuleIsSignature"
Option Explicit
Option Compare Text

Const commentedOut As String = "'Public Property Get HelloThere() As String"
Const allOneWord As String = "If IsObject(item) And Not compareByDefaultProperty Then"
Const legitProperty As String = "Public Property Get HelloThere() As String"
Const legitPropertyWithSpace As String = " Public Property Get HelloThere() as string"
Const legitPropertyWithTab As String = "    Public Property Get HelloThere() As String"

Private Function IsSignature(line As String) As Boolean
    IsSignature = (line Like "[!']* Property *")
End Function

Private Sub Test()
    Debug.Assert Not IsSignature(commentedOut)
    Debug.Assert Not IsSignature(allOneWord)
    Debug.Assert IsSignature(legitProperty)
    Debug.Assert IsSignature(legitPropertyWithSpace)
    Debug.Assert IsSignature(legitPropertyWithTab)
End Sub
