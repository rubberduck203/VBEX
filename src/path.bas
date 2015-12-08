Attribute VB_Name = "path"
Option Explicit
'
' path
' ====
'
' Common Path Manipulations for VBEX
'
' Copyright (c) 2014 Philip Wales
' This file (path.bas) is distributed under the GPL-3.0 license.
'


'
' Constants
' ---------
'
Public Const EXTSEP As String = "."
Public Const PARDIR As String = ".."
Public Const CURDIR As String = "."
Public Const SEP As String = "\" ' "/" for UNIX if you ever run VBA on UNIX...
Public Const PATHSEP As String = ";" ' not used...
'
' Path Manipulations
' ------------------
'
''
' Returns the base name of a path, either the lowest folder or file
' Note! that `suffix` will be removed from the end regardless if its an actual filename
' extension or not.
Public Function BaseName(ByVal filePath As String, _
        Optional ByVal suffix As String) As String

    Dim pathSplit As Variant
    pathSplit = Split(filePath, SEP)
    
    BaseName = pathSplit(UBound(pathSplit))
    
    If suffix <> vbNullString Then
    
        Dim baseLength As Integer
        baseLength = Len(BaseName) - Len(suffix)
        
        ' replace suffix with nothing and only look for suffix the end of the string
        BaseName = Left$(BaseName, baseLength) & Replace$(BaseName, suffix, "", baseLength + 1)
        
    End If
    
End Function
''
' Returns the path of the parent folder. This is the opposite of `BaseName`.
Public Function RootName(ByVal Path As String) As String

    RootName = ParentDir(Path, 1)
    
End Function
''
'
Public Function ParentDir(ByVal somePath As String, _
        ByVal parentHeight As Integer) As String
    
    Dim splitPath As Variant
    splitPath = Split(somePath, SEP)
    
    Dim parentCount As Integer
    parentCount = UBound(splitPath) - parentHeight
    
    If parentCount > 0 Then

        ReDim Preserve splitPath(LBound(splitPath) To parentCount)
        
    End If
     
    ParentDir = Join(splitPath, SEP)
   
End Function
''
' Returns the file extension of the file.
' path.ext -> .ext
' path ->
' path.bad.ext -> .ext
Public Function Ext(ByVal filePath As String) As String

    Dim base As String
    base = BaseName(filePath)
    
    If InStr(base, EXTSEP) Then
    
        Dim fsplit As Variant
        fsplit = Split(base, EXTSEP)
        
        Ext = EXTSEP & fsplit(UBound(fsplit))
        
    End If
    
End Function
''
' Removes trailing SEP from path
Public Function RTrimSep(ByVal Path As String) As String

    If right$(Path, 1) = SEP Then
        ' ends with SEP return all but end
        RTrimSep = Left$(Path, Len(Path) - 1)
    Else
        RTrimSep = Path
    End If
    
End Function
''
' safely join two strings to form a path, inserting `SEP` if needed.
Public Function JoinPath(ByVal rootPath As String, ByVal filePath As String) As String

    JoinPath = RTrimSep(rootPath) & SEP & filePath
    
End Function
''
' Inserts `toAppend` in behind of the base name of string `filePath` but in
' front of the extension
Public Function Append(ByVal filePath As String, ByVal toAppend As String) As String

    Dim fileExt As String
    fileExt = Ext(filePath)
    
    Dim Root As String
    Root = RootName(filePath)
    
    Dim base As String
    base = BaseName(filePath, suffix:=fileExt)
    
    Dim newName As String
    newName = base & toAppend & fileExt
    
    Append = JoinPath(Root, newName)
                     
End Function
''
' Inserts `toPrepend` in front of the base name of string `filePath`
' root/name.ext -> prepended -> root/prependedname.ext
Public Function Prepend(ByVal filePath As String, ByVal toPrepend As String) As String
    
    Prepend = JoinPath(RootName(filePath), toPrepend & BaseName(filePath))

End Function
''
' Replaces current extension of `filePath` with `newExt`
Public Function ChangeExt(ByVal filePath As String, ByVal newExt As String) As String
    
    Dim currentExt As String
    currentExt = Ext(filePath)
    
    Dim baseLength As String
    baseLength = Len(filePath) - Len(currentExt)
    
    ' ".ext" or "ext" -> "ext"
    newExt = Replace$(newExt, EXTSEP, vbNullString, 1, 1)

    ChangeExt = Left$(filePath, baseLength) & EXTSEP & newExt
    
End Function
''
' Returns if the filePath contains a "?" or a "*"
Public Function IsPattern(ByVal filePath As String) As Boolean
    IsPattern = (InStr(1, filePath, "?") + InStr(1, filePath, "*") <> 0)
End Function
''
' Finds the longest filePath in pattern that is not a pattern.
Public Function LongestRoot(ByVal pattern As String) As String
    
    Dim charPos As Integer
    charPos = InStr(1, pattern, "?") - 1

    Dim wildPos As Integer
    wildPos = InStr(1, pattern, "*") - 1

    Dim firstPatternPos As Integer
    If wildPos < 0 And charPos < 0 Then ' not a pattern
        firstPatternPos = Len(pattern)
    ElseIf wildPos < 0 Then
        firstPatternPos = charPos
    ElseIf charPos < 0 Then
        firstPatternPos = wildPos
    Else
        firstPatternPos = srch.Min(wildPos, charPos)
    End If
    
    LongestRoot = RootName(Left$(pattern, firstPatternPos))
    
End Function
