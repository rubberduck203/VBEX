Attribute VB_Name = "shutilB"
Option Explicit
'
' shutilB
' =======
'
' Advanced filesystem operations for VBA. Routines of this variant will return `true` if
' operation is successful or `false` if failed.
'
' Copyright (c) 2014 Philip Wales
' This file (shutilB.bas) is distributed under the MIT license.
' Obtain a copy of the license here: http://opensource.org/licenses/MIT
'
' Scripting.FileSystemObject is slow and unstable since it relies on sending
' signals to ActiveX objects across the system.  This module only uses built-in
' functions of Visual Basic, such as `Dir`, `Kill`, `Name`, etc.

'
'
' File System Modifications
' -------------------------
'
'
Function Move(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
              
    Dim check As Boolean
    On Error GoTo ErrHandler

    DestIsFolderFeature dest, src
    
    If createParent Then CreateRootPath dest
    Name src As dest
    check = fsview.Exists(dest)
    
CleanExit:
    Move = check
    Exit Function
  
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Function Rename(ByVal aPath As String, ByVal newName As String) As Boolean
    
    Debug.Assert path.BaseName(newName) = newName
    
    Rename = Move(aPath, path.JoinPath(path.RootName(aPath), newName))

End Function
Function Remove(ByVal filePath As String) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler
    
    Kill filePath
    check = (Not fsview.FileExists(filePath))
    
CleanExit:
    Remove = check
    Exit Function

ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit

End Function
Function MakeDir(ByVal filePath As String, _
        Optional createParent As Boolean = False) As Boolean
                
    Dim check As Boolean
    On Error GoTo ErrHandler
        
    If createParent Then CreateRootPath filePath
    MkDir filePath
    check = fsview.FolderExists(filePath)
    
CleanExit:
    MakeDir = check
    Exit Function
    
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Function CopyFile(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler
    
    DestIsFolderFeature dest, src
    
    If fsview.FileExists(dest) Then GoTo CleanExit:
    
    If createParent Then CreateRootPath dest
    FileCopy src, dest
    check = fsview.FileExists(dest)

CleanExit:
    CopyFile = check
    Exit Function
    
ErrHandler:
    Err.Clear
    Debug.Assert (Not check)
    Resume CleanExit
    
End Function
Private Function CreateRootPath(ByVal aPath As String) As Boolean

    Dim parentFolder As String
    parentFolder = path.RootName(aPath)
    
    If Not fsview.FolderExists(parentFolder) Then
    
        CreateRootPath = MakeDir(parentFolder, createParent:=True)
        
    End If
    
End Function
Private Sub DestIsFolderFeature(ByRef dest As String, ByVal src As String)

    If right$(dest, 1) = path.SEP Or fsview.FolderExists(dest) Then
        dest = path.JoinPath(dest, path.BaseName(src))
    End If
    
End Sub
