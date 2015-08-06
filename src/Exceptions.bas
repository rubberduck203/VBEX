Attribute VB_Name = "Exceptions"
Option Explicit
Public Enum vbErrorNums
    TYPE_ERROR = 13
    OBJECT_REQUIRED = 424
End Enum
Public Enum exErrorNums
    TYPE_ERROR = 13
    UNIMPLEMENTED = 1
    VALUE_ERROR = 666
End Enum
'
' Exceptions
' ==========
'
Public Sub BubbleError(ByVal raiser, ByVal method As String, _
        ByVal e As ErrObject)

    Dim trace As String
    trace = MakeDescription(raiser, method, e.description)

    Err.Raise e.Number, e.source, trace, e.HelpFile, e.HelpContext

End Sub
Public Sub IndexError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
    
    NotImplementedError "Exceptions", "IndexError"
    
End Sub
Public Sub KeyError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    NotImplementedError "Exceptions", "KeyError"
    
End Sub
Public Sub NotImplementedError(ByVal raiser, ByVal method As String)

    Dim source As String
    source = MakeSource(raiser, method)
    
    Dim msg As String
    msg = source & " Not implemented."

    Err.Raise exErrorNums.UNIMPLEMENTED, description:=MakeDescription(raiser, method, msg)
    
End Sub
Public Sub OSError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    NotImplementedError "Exceptions", "OSError"
    
End Sub
Public Sub TypeError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    Err.Raise exErrorNums.TYPE_ERROR, description:=MakeDescription(raiser, method, msg)
    
End Sub
Public Sub ValueError(ByVal raiser, ByVal method As String, _
        Optional ByVal msg As String)
        
    Err.Raise exErrorNums.VALUE_ERROR, description:=MakeDescription(raiser, method, msg)
    
End Sub
'
' Private Methods
' ---------------
'
Private Function MakeDescription(ByVal raiser, ByVal method As String, _
        ByVal msg As String) As String
    
    MakeDescription = AddTrace(MakeSource(raiser, method), msg)
    
End Function
Private Function MakeSource(ByVal raiser, ByVal method As String) As String

    Dim result As String
    If IsObject(raiser) Then
        result = TypeName(raiser) & "." & method
    Else
        result = raiser & "." & method
    End If

    MakeSource = result

End Function
Private Function AddTrace(ByVal source As String, _
        ByVal description As String) As String

    AddTrace = source & " >> " & description

End Function
