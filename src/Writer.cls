VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' File Writer
' ===========
'
' author: Matt's Mug
' url: http://codereview.stackexchange.com/questions/52306/file-reader-writer-text
'
'
Option Explicit
Implements IOutput

'
' Private Members
' ---------------
'
Private Const ERROR_BASE As Long = &HFADE
Private pFilePath As String
Private pFileNumber As Long
'
' Public Memebers
' ---------------
'
'expose raised errors to clients:
Public Enum FileWriterError
    FileNotOpened = vbObjectError + ERROR_BASE + 42
    FileAlreadyOpened
    FileDoesNotExist
    FileAlreadyExists
End Enum
'
' Private Methods
' ---------------
'
Private Sub Class_Terminate()

    CloseFile
    
End Sub
'
' ### Writing
'
Private Sub AppendData(ByVal data As String, _
        ByVal withLineFeed As Boolean, ByVal method As String)
On Error GoTo ErrHandler

    If pFileNumber = 0 Then
        OnFileNotOpenedError method
    End If

    If withLineFeed Then
        Print #pFileNumber, data
    Else
        Print #pFileNumber, data;
    End If

CleanExit:
    Exit Sub

ErrHandler:

    'handle "52: Bad file name or number" by raising a FileWriterError.FileNotOpened instead:
    If Err.Number = 52 Then OnFileNotOpenedError method

    'close file it *any* error occurs writing to it:
    CloseFile

    'bubble up all errors
    Err.Raise Err.Number, Err.source, Err.Description
End Sub
'
' ### Errors
'
Private Function GetErrorSource(ByVal method As String) As String

    GetErrorSource = TypeName(Me) & "." & method
    
End Function
Private Sub OnFileNotOpenedError(ByVal method As String)

    Err.Raise FileWriterError.FileNotOpened, GetErrorSource(method), _
        "File #" & pFileNumber & "(" & pFilePath & ") was unexpectedly closed."
        
End Sub
Private Sub OnFileAlreadyOpenedError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileAlreadyOpened, GetErrorSource(method), _
        "File '" & fileName & _
        "' cannot be opened with this instance at this point. A file is already opened."
        
End Sub
Private Sub OnFileDoesNotExistError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileDoesNotExist, GetErrorSource(method), _
        "File '" & fileName & "' cannot be opened as it does not exist."
        
End Sub
Private Sub OnFileAlreadyExistsError(ByVal method As String, ByVal fileName As String)

    Err.Raise FileWriterError.FileAlreadyExists, GetErrorSource(method), _
        "File '" & fileName & "' cannot be created as it already exists."
        
End Sub
'
' Friend Methods
' --------------
'
Friend Sub pOpenFile(ByVal filePath As String, _
        Optional ByVal overwrite As Boolean = True)

    If pFileNumber <> 0 Then
        OnFileAlreadyOpenedError "OpenFile", filePath
    End If
    
    If Not fsview.FileExists(filePath) Then
        OnFileDoesNotExistError "OpenFile", filePath
    End If
    
    pFileNumber = FreeFile
    pFilePath = filePath

    If overwrite Then
        Open filePath For Output As #pFileNumber
    Else
        Open filePath For Append As #pFileNumber
    End If

Exit Sub

ErrHandler:
    Err.Raise Err.Number, GetErrorSource("OpenFile"), Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub
Friend Sub pNewFile(ByVal filePath As String)

    If pFileNumber <> 0 Then
        OnFileAlreadyOpenedError "NewFile", filePath
    End If
    
    If fsview.FileExists(filePath) Then
        OnFileAlreadyExistsError "NewFile", filePath
    End If
    
    pFileNumber = FreeFile
    pFilePath = filePath

    If Not fsview.FileExists(filePath) Then
        Open filePath For Output As #pFileNumber
    End If

Exit Sub

ErrHandler:
    Err.Raise Err.Number, GetErrorSource("NewFile"), Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub
'
' Constructor
' -----------
'
Public Function OpenFile(ByVal filePath As String, _
        Optional ByVal overwrite As Boolean = True) As Writer

    Dim result As New Writer
    result.pOpenFile filePath, overwrite:=overwrite
    Set OpenFile = result
    
End Function
Public Function NewFile(ByVal filePath As String) As Writer

    Dim result As New Writer
    result.pNewFile filePath
    Set NewFile = result
    
End Function
'
' Public Methods
' --------------
'

Public Sub CloseFile()
On Error GoTo ErrHandler

    Close #pFileNumber
    pFileNumber = 0

Exit Sub

ErrHandler:
    Err.Raise Err.Number, GetErrorSource("CloseFile"), Err.Description, _
        Err.HelpFile, Err.HelpContext
        
End Sub
'
' Interfaces
' ----------
'
' ### IOutput
'
Public Sub PrintLine(ByVal data As Variant)
    
    AppendData cast.ToString(data), withLineFeed:=True, method:="PrintLine"
    
End Sub
Public Sub Puts(ByVal data As Variant)

    AppendData cast.ToString(data), withLineFeed:=False, method:="Puts"
    
End Sub
Public Sub IOutput_PrintLine(ByVal data As Variant)

    PrintLine data

End Sub
Public Sub IOutput_Puts(ByVal data As Variant)

    Puts data

End Sub


