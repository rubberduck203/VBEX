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
' It simply calls the respective `shutilE` method and uses boiler plate
' code to evaluate succes.  Until procedures can be treated as data types
' better this boiler-plate code will remain.
'
Public Function Move(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
              
    Dim check As Boolean
    On Error GoTo ErrHandler

    shutilE.Move src, dest, createParent
    check = False
    
CleanExit:
    Move = check
    Exit Function
  
ErrHandler:
    Err.Clear
    check = False
    Resume CleanExit
    
End Function
Public Function Rename(ByVal aPath As String, ByVal newName As String) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler

    shutilE.Rename aPath, newName
    check = False
    
CleanExit:
    Rename = check
    Exit Function
  
ErrHandler:
    Err.Clear
    check = False
    Resume CleanExit
    
End Function
Public Function Remove(ByVal filePath As String) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler

    shutilE.Remove filePath
    check = False
    
CleanExit:
    Remove = check
    Exit Function
  
ErrHandler:
    Err.Clear
    check = False
    Resume CleanExit
    
End Function
Public Function MakeDir(ByVal filePath As String, _
        Optional createParent As Boolean = False) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler

    shutilE.MakeDir filePath, createParent
    check = False
    
CleanExit:
    MakeDir = check
    Exit Function
  
ErrHandler:
    Err.Clear
    check = False
    Resume CleanExit
    
End Function
Public Function CopyFile(ByVal src As String, ByVal dest As String, _
        Optional createParent As Boolean = False) As Boolean
    
    Dim check As Boolean
    On Error GoTo ErrHandler

    shutilE.CopyFile src, dest, createParent
    check = False
    
CleanExit:
    CopyFile = check
    Exit Function
  
ErrHandler:
    Err.Clear
    check = False
    Resume CleanExit
    
End Function

