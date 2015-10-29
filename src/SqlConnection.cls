VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'
' SqlConnection
' =============
'
' Establishes an ADODB connection to a source and provides a rich interface to
' that source.
'
Implements IShowable
'
' Private Member
' --------------
'
Private pConnection As ADODB.Connection
'
' Error Values
' ------------
'
Public Enum SqlConnectionErrors
    NO_RECORDS = 770
    TOO_MANY_RECORDS
    TOO_MANY_FIELDS
End Enum
'
' Constructor
' -----------
'
' ### Constructors
'
Public Function Make(ByVal connectString As String) As SqlConnection

    Dim result As New SqlConnection
    result.OpenConnection connectString
    Set Make = result
    
End Function
Public Function ToAccess(ByVal filePath As String) As SqlConnection
    
    Dim connectString As String
    connectString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & filePath & ";" & _
        "Jet OLEDB:Engine Type=5;" & _
        "Persist Security Info=False;"

    Set ToAccess = Make(connectString)

End Function
Public Function ToExcel(ByVal filePath As String) As SqlConnection

    Dim connectString As String
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & filePath & ";" & _
        "Extended Properties=Excel 8.0;"

    Set ToExcel = Make(connectString)

End Function
Public Function ToMySql(ByVal server As String, ByVal db As String, _
        ByVal username As String, ByVal passwd As String) As SqlConnection
        
    Dim connectTemplate As String
    connectTemplate = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER={0}; DATABASE={1}; UID={2};PASSWORD={3}; OPTION=3"

    Dim connect As String
    connect = Str.Format(connectTemplate, server, db, username, passwd)
    
    Set ToMySql = Make(connect)
    
End Function
'
' ### Class Events
'
Private Sub Class_Initialize()
                             
    Set pConnection = New ADODB.Connection

End Sub
Private Sub Class_Terminate()

    If pConnection.State Then pConnection.Close

End Sub
'
' ### Friend Properties
'
Friend Sub OpenConnection(ByVal connectStr As String)

    pConnection.Open connectStr

End Sub
Friend Property Get DbConnection() As ADODB.Connection

    Set DbConnection = pConnection

End Property
'
' Interfaces
' ----------
'
' ### IShowable
'
Private Function IShowable_Show() As String

    IShowable_Show = defshow.ParamShowableObject(Me, pConnection.ConnectionString)

End Function
'
'#######################################################################################
'
' Returning SQLCommands
'
' Custom SQL
Public Function MakeCommand(ByVal sql As String) As SqlCommand

    Set MakeCommand = SqlCommand.Make(MakeAdoCommand(sql))

End Function
Public Function MakeAdoCommand(ByVal sql As String) As ADODB.Command
    
    Dim result As New ADODB.Command
    result.ActiveConnection = DbConnection
    result.CommandText = sql
    Set MakeAdoCommand = result

End Function
' Stored SQL
Public Function GetCommand(ByVal sqlName As String) As SqlCommand

     Set GetCommand = SqlCommand.Make(GetAdoCommand(sqlName))

End Function
Public Function GetAdoCommand(ByVal sqlName As String) As ADODB.Command

    Dim result As ADODB.Command
    Set result = MakeAdoCommand(sqlName)
    result.CommandType = adCmdStoredProc
    Set GetAdoCommand = result

End Function
'
' ### Append safely
'
Public Function SafeAdd(ByVal getArgs As Dict, ByVal getQuerysName As String, _
        ByVal addArgs As Dict, ByVal addQuerysName As String) As Long
    
    Dim finder As SqlCommand
    Set finder = GetCommand(getQuerysName)
    
    Dim autoId As Maybe
    Set autoId = finder.ExecuteQueryOnArray(getArgs.Values.ToArray).SingleResult
    
    If autoId.IsNone Then
    
        Dim checkRow As Long
        checkRow = GetCommand(addQuerysName).ExecuteUpdateOnArray(addArgs.Values.ToArray)
        Debug.Assert checkRow = 1
        Set autoId = finder.ExecuteQueryOnArray(getArgs.Values.ToArray).SingleResult
        
    End If
    
    On Error GoTo ErrHandler
    SafeAdd = autoId
    
Exit Function
ErrHandler:
    Dim showArgs As Dict
    Set showArgs = Dict.Create( _
        Assoc.Make("getArgs", getArgs), _
        Assoc.Make("getQuerysName", getQuerysName), _
        Assoc.Make("addArgs", addArgs), _
        Assoc.Make("addQuerysName", addQuerysName) _
    )
    
    Exceptions.ValueError Me, "SafeAdd", "Failed to get autoId with " & showArgs.Show
    
End Function
Public Function SimpleSafeAdd(ByVal arg As Assoc, ByVal getQuerysName As String, _
        ByVal addQuerysName As String) As Long
    
    Dim unifiedArgs As Dict
    Set unifiedArgs = Dict.Create(arg)
    
    SimpleSafeAdd = SafeAdd(unifiedArgs, getQuerysName, unifiedArgs, addQuerysName)
    
End Function

