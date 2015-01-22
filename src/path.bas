Attribute VB_Name = "path"
Option Explicit
'
' path
' ====
'
' Common Path Manipulations for VBEX
'
' Copyright (c) 2014 Philip Wales
' This file (path.bas) is distributed under the MIT license.
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
' root/name.ext -> name.ext
' name.ext -> name.ext
' root/ ->
' root/name+suffix -> suffix -> name
' "root/name.ext" -> ".ext" -> "name"
' "root/name.ext" -> "ext" -> "name."  !
Public Function BaseName(ByVal file_path As String, Optional ByVal suffix As String) As String

    Dim path_split As Variant
    path_split = Split(file_path, SEP)
    
    BaseName = path_split(UBound(path_split))
    
    If suffix <> vbNullString Then
    
        Dim base_length As Integer
        base_length = Len(BaseName) - Len(suffix)
        
        ' replace suffix with nothing and only look for suffix the end of the string
        BaseName = Left$(BaseName, base_length) & _
                   Replace$(BaseName, suffix, "", base_length + 1)
                   
    End If
    
End Function
''
' Returns the path of the parent folder. This is the opposite of `BaseName`.
' r/o/o/t/name -> r/o/o/t
' r/o/o/t/ -> r/o/o/t
' name ->
Public Function RootName(ByVal path As String) As String

    RootName = ParentDir(path, 1)
    
End Function
''
' path -> 0 -> path
' path/ -> 1 -> path
' root/name -> 1 -> root ! `RootName`
' E/D/C/B/A/name -> 2 -> E/D/C/B
' E/D/C/B/A/name -> 3 -> E/D/C
' E/D/C/B/A/name -> 5 -> E
' E/D/C/B/A/name -> 6 ->
' E/D/C/B/A/name -> 7 ->
' ...
Public Function ParentDir(ByVal path As String, _
                   ByVal parent_height As Integer) As String
    
    Dim split_path As Variant
    split_path = Split(path, SEP)
    
    Dim parent_count As Integer
    parent_count = UBound(split_path) - parent_height
    
    If parent_count > 0 Then

        ReDim Preserve split_path(LBound(split_path) To parent_count)
        
    End If
     
    ParentDir = Join(split_path, SEP)
   
End Function
''
' Returns the file extension of the file.
' path.ext -> .ext
' path ->
' path.bad.ext -> .ext
Public Function Ext(ByVal file_path As String) As String

    Dim base_name As String
    base_name = BaseName(file_path)
    
    If InStr(base_name, EXTSEP) Then
    
        Dim fsplit As Variant
        fsplit = Split(base_name, EXTSEP)
        
        Ext = EXTSEP & fsplit(UBound(fsplit))
        
    End If
    
End Function
''
' Removes trailing SEP from path
' path/ -> path
' path -> path
' /path -> /path
Public Function RTrimSep(ByVal path As String) As String

    If right$(path, 1) = SEP Then
        ' ends with SEP return all but end
        RTrimSep = Left$(path, Len(path) - 1)
        
    Else
        RTrimSep = path
        
    End If
    
End Function
''
' safely join two strings to form a path, inserting `SEP` if needed.
' root/ -> base -> root/base
' root -> base -> root/base
' root -> /base -> root//base ! BAD BAD BAD
Public Function JoinPath(ByVal root_path As String, ByVal file_path As String) As String

    JoinPath = RTrimSep(root_path) & SEP & file_path
    
End Function
''
' Inserts `to_append` in behind of the base name of string `file_path` but in
' front of the extension
' root/name.ext -> appended -> root/nameappended.ext
Public Function Append(ByVal file_path As String, ByVal to_append As String) As String

    Dim file_ext As String
    file_ext = Ext(file_path)
    
    Append = JoinPath(RootName(file_path), _
                   BaseName(file_path, suffix:=file_ext) & _
                   to_append & file_ext)
                     
End Function
''
' Inserts `to_prepend` in front of the base name of string `file_path`
' root/name.ext -> prepended -> root/prependedname.ext
Public Function Prepend(ByVal file_path As String, ByVal to_prepend As String) As String
    
    Prepend = JoinPath(RootName(file_path), to_prepend & BaseName(file_path))

End Function
''
' Replaces current extension of `file_path` with `new_ext`
' path.old -> new -> path.new
' path.old -> .new -> path.new
' path -> new -> path.new
' path.bad.old -> new -> path.bad.new
Public Function ChangeExt(ByVal file_path As String, ByVal new_ext As String) As String
    
    Dim current_ext As String
    current_ext = Ext(file_path)
    
    Dim base_length As String
    base_length = Len(file_path) - Len(current_ext)
    
    ' ".ext" or "ext" -> "ext"
    new_ext = Replace$(new_ext, EXTSEP, vbNullString, 1, 1)

    ChangeExt = Left$(file_path, base_length) & EXTSEP & new_ext
    
End Function
''
' Returns if the path contains a "?" or a "*"
Public Function IsPattern(ByVal path As String) As Boolean
    IsPattern = (InStr(1, path, "?") + InStr(1, path, "*") <> 0)
End Function
''
' Finds the longest path in pattern that is not a pattern.
Public Function LongestRoot(ByVal pattern As String) As String
    
    Dim charPos As Integer
    charPos = InStr(1, pattern, "?") - 1
    If charPos < 0 Then charPos = Len(pattern)
    
    Dim wildPos As Integer
    wildPos = InStr(1, pattern, "*") - 1
    If wildPos < 0 Then wildPos = Len(pattern)

    LongestRoot = RootName(Left$(pattern, IIf(charPos <= wildPos, charPos, wildPos)))
    
End Function
