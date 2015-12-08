Attribute VB_Name = "shutilE"
Option Explicit
'
' shutilE
' =======
'
' Advanced filesystem operations for VBA. This variant raises errors if
' attempted operation fails.
'
' Copyright (c) 2014 Philip Wales
' This file (shutilE.bas) is distributed under the GPL-3.0 license.
' Obtain a copy of the license here: http://opensource.org/licenses/GPL-3.0
'
' Scripting.FileSystemObject is slow and unstable since it relies on sending
' signals to ActiveX objects across the system.  This module only uses built-in
' functions of Visual Basic, such as `Dir`, `Kill`, `Name`, etc.
Public Enum ShutilErrors
    overWRiteRefusal
    failedDestroy
    failedCreate
End Enum
'
'
' File System Modifications
' -------------------------
'
'
Public Sub Move(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False)

    On Error GoTo ErrHandler

    DestIsFolderFeature dest, src
    
    If createParent Then CreateRootPath dest
    
    Name src As dest
    
    If Not fsview.Exists(dest) Then
        OnFailedCreateError "Move", "Name As"
    End If
    
    If fsview.Exists(src) Then
        OnFailedDestroyError "Move", "Name As"
    End If
    
CleanExit:
    Exit Sub
  
ErrHandler:
    Select Case Err.Number
    Case Else
        ReRaiseError Err
    End Select

End Sub
Public Sub Rename(ByVal src As String, ByVal newName As String)

    On Error GoTo ErrHandler
    
    Debug.Assert newName = Path.BaseName(newName)
    
    Dim Root As String
    Root = RootName(src)
    
    Dim dest As String
    dest = Path.JoinPath(Root, newName)
    
    Move src, dest
    
CleanExit:
    Exit Sub
  
ErrHandler:
    Select Case Err.Number
    Case Else
        ReRaiseError Err
    End Select

End Sub
Public Sub Remove(ByVal aPath As String)
    On Error GoTo ErrHandler
    
    Kill aPath
    
    If fsview.Exists(aPath) Then
        OnFailedDestroyError "Remove", "Kill"
    End If
    
CleanExit:
    Exit Sub

ErrHandler:
    Select Case Err.Number
    Case Else
        ReRaiseError Err
    End Select
    
End Sub
Public Sub MakeDir(ByVal folderPath As String, Optional ByVal createParent As Boolean = False)

    Dim check As Boolean
    On Error GoTo ErrHandler
        
    If createParent Then CreateRootPath folderPath
    MkDir folderPath
    
    If Not fsview.FolderExists(folderPath) Then
        OnFailedCreateError "MakeDir", "MkDir"
    End If
    
CleanExit:
    Exit Sub
    
ErrHandler:
    Select Case Err.Number
    Case Else
        ReRaiseError Err
    End Select
    
End Sub
Public Sub CopyFile(ByVal src As String, ByVal dest As String, _
      Optional createParent As Boolean = False)
    
    On Error GoTo ErrHandler
    
    DestIsFolderFeature dest, src
    
    If fsview.FileExists(dest) Then
        OnNoOverwriteError "CopyFile"
    End If
    
    If createParent Then CreateRootPath dest
    FileCopy src, dest
    
    If Not fsview.FileExists(dest) Then
        OnFailedCreateError "CopyFile", "FileCopy"
    End If

CleanExit:
   Exit Sub

ErrHandler:
    Select Case Err.Number
    Case Else
       ReRaiseError Err
    End Select
    
End Sub
Private Sub CreateRootPath(ByVal aPath As String)
    
    Dim parentFolder As String
    parentFolder = Path.RootName(aPath)
    
    If Not fsview.FolderExists(parentFolder) Then
        MakeDir parentFolder, createParent:=True
    End If
    
End Sub
Private Sub DestIsFolderFeature(ByRef dest As String, _
        ByVal src As String)
    
    If right$(dest, 1) = Path.SEP Or fsview.FolderExists(dest) Then
        dest = Path.JoinPath(dest, Path.BaseName(src))
    End If
    
End Sub
'
' ### Custom Error Messages
'
Private Sub ReRaiseError(ByRef e As ErrObject)

    Err.Raise e.Number, e.source, e.description, e.HelpFile, e.HelpContext
    
End Sub
Private Sub OnFailedCreateError(ByVal method As String, ByVal operation As String)

    Err.Raise ShutilErrors.failedCreate, method, _
        "Destination does not exist after errorless `" & operation & "`"
        
End Sub
Private Sub OnFailedDestroyError(ByVal method As String, ByVal operation As String)

    Err.Raise ShutilErrors.failedDestroy, method, _
        "Destination still exists after errorless `" & operation & "`"
    
End Sub
Private Sub OnNoOverwriteError(ByVal method As String)

    Err.Raise ShutilErrors.overWRiteRefusal, method, _
        "Will not overwrite file at destination.  Remove it first if desired."
    
End Sub
