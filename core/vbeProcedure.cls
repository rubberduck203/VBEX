VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' requires Microsoft Visual Basic for Applications Extensibility 5.3 library

' *******************************************************************
' * Class: vbeProcedure                                             *
' * Fills in some gaps in the VBIDE library                         *
' * Author: Christopher J. McClellan                                *
' * Date Created: 5/30/14                                           *
' * Published under Creative Commons Attribution-Share Alike        *
' * http://creativecommons.org/licenses/by-sa/3.0/                  *
' * You are free to change, distribute, and pretty much do          *
' *     whatever you like with the code, but you must give credit   *
' *     to the original author and publish any derivitive of this   *
' *     code under the same license.                                *
' *******************************************************************

Option Explicit

' error handling values
Private Const BaseErrorNum As Long = 3500

Public Enum vbeProcedureError
        vbeObjectNotIntializedError = vbObjectError + BaseErrorNum
        vbeReadOnlyPropertyError
        vbeInvalidArgError
End Enum

Public Enum MemberType
    mt_PropertyGetter
    mt_PropertyLetter
    mt_PropertySetter
    mt_Function
    mt_Sub
End Enum

Public Enum MemberAccessibility
    ma_Public
    ma_Private
    ma_Friend
End Enum

Private Const ObjectNotIntializedMsg = "Object Not Initialized"
Private Const ReadOnlyPropertyMsg = "Property is Read-Only after initialization"

' exposed property variables
Private Type TVbeProcedure
    ParentModule As CodeModule
    Name As String
    procKind As vbext_ProcKind
End Type

Private this As TVbeProcedure

' truly private property variables
Private isNameSet As Boolean
Private isParentModSet As Boolean

Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    If isNameSet Then
        Name = this.Name
    Else
        RaiseObjectNotIntializedError
    End If
End Property

Public Property Let Name(ByVal vNewValue As String)
    If Not isNameSet Then
        If vNewValue = vbNullString Then
            RaiseInvalidArgError "Name", "The Name property can not be set to an empty string."
        End If
        this.Name = vNewValue
        isNameSet = True
    Else
        RaiseReadOnlyPropertyError
    End If
End Property

Public Property Get ParentModule() As CodeModule
Attribute ParentModule.VB_Description = "Gets or Sets the Parent VBIDE.CodeModule"
    If isParentModSet Then
        Set ParentModule = this.ParentModule
    Else
        RaiseObjectNotIntializedError
    End If
End Property

Public Property Let ParentModule(ByRef vNewValue As CodeModule)
' Object assignments should use Set, but that forces Initialize() through the Getter, raising ObjectNotInitialized
    If Not isParentModSet Then
        Set this.ParentModule = vNewValue
        isParentModSet = True
    Else
        RaiseReadOnlyPropertyError
    End If
End Property

Public Property Get procKind() As vbext_ProcKind
    procKind = this.procKind
End Property

Public Property Get StartLine() As Long
    ValidateIsInitialized
    StartLine = Me.ParentModule.ProcStartLine(Me.Name, this.procKind)
End Property

Public Property Get EndLine() As Long
    ValidateIsInitialized
    EndLine = Me.StartLine + Me.CountOfLines
End Property

Public Property Get CountOfLines() As Long
    ValidateIsInitialized
    CountOfLines = Me.ParentModule.ProcCountLines(Me.Name, this.procKind)
End Property

Public Sub Initialize(Name As String, codeMod As CodeModule, procKind As vbext_ProcKind)
    Me.Name = Name
    Me.ParentModule = codeMod
    this.procKind = procKind
End Sub

Public Property Get Lines() As String
    ValidateIsInitialized
    Lines = Me.ParentModule.Lines(Me.StartLine, Me.CountOfLines)
End Property

Public Property Get Signature() As String
' @Mat's Mug [http://codereview.stackexchange.com/users/23788/mats-mug] wrote this.
    Dim code() As String
    code = Split(Me.ParentModule.Lines(Me.StartLine, Me.CountOfLines), vbNewLine)

    Dim i As Long
    For i = 0 To UBound(code)
        If code(i) <> vbNullString And Left(Trim(code(i)), 1) <> "'" Then
            Signature = code(i)
            Exit Property
        End If
    Next
End Property

'TODO: Property Body
'Public Property Get Body() As String

'End Property

Public Property Get ModuleMemberType() As MemberType
' @Mat's Mug [http://codereview.stackexchange.com/users/23788/mats-mug] wrote this.
    Dim result As MemberType
    Dim code() As String
    code = Split(Trim(Signature), " ")
    
    Dim modifier As String
    modifier = code(0)
    
    Dim mType As String, mPropType As String
    If modifier = "Property" Or modifier = "Function" Or modifier = "Sub" Then
        mType = modifier
        mPropType = code(1)
    Else
        mType = code(1)
        mPropType = code(2)
    End If
    
    Select Case mType
        Case "Property"
            
            If mPropType = "Get" Then
                ModuleMemberType = mt_PropertyGetter
            ElseIf mPropType = "Let" Then
                ModuleMemberType = mt_PropertyLetter
            ElseIf mPropType = "Set" Then
                ModuleMemberType = mt_PropertySetter
            Else
                Const InvalidProcedureCallOrArguement As Long = 5
                Err.Raise InvalidProcedureCallOrArguement
            End If
            
        Case "Function"
            ModuleMemberType = mt_Function
            
        Case "Sub"
            ModuleMemberType = mt_Sub
           
    End Select
    
End Property

Property Get Accessibility() As MemberAccessibility
' @Mat's Mug [http://codereview.stackexchange.com/users/23788/mats-mug] wrote this.
    Dim code() As String
    code() = Split(Trim(Signature), " ")
    
    Dim modifier As String
    modifier = code(0)
    
    If modifier = "Property" Or modifier = "Function" Or modifier = "Sub" Then modifier = "Public"
    
    Select Case modifier
        Case "Public"
            Accessibility = ma_Public
        Case "Private"
            Accessibility = ma_Private
        Case "Friend"
            Accessibility = ma_Friend
        Case Else
            Err.Raise 5
    End Select
    
End Property

' TODO: Property ReturnType; get the properties Type
' TODO: "Create" or "Append" sub
' TODO: "Insert" Sub
' TODO: Sort function

Private Sub RaiseObjectNotIntializedError()
    Err.Raise vbeProcedureError.vbeObjectNotIntializedError, GetErrorSource, ObjectNotIntializedMsg
End Sub

Private Sub RaiseReadOnlyPropertyError()
    Err.Raise vbeProcedureError.vbeReadOnlyPropertyError, GetErrorSource, ReadOnlyPropertyMsg
End Sub

Private Sub RaiseInvalidArgError(propertyName As String, Optional additonalInfo As String = vbNullString)
    Dim message As String
    message = "Invalid Argument" & vbCrLf & "Property: " & propertyName
        
    If additonalInfo = vbNullString Then
        Err.Raise vbeProcedureError.vbeInvalidArgError, GetErrorSource, message
    Else
        Err.Raise vbeProcedureError.vbeInvalidArgError, GetErrorSource, message & vbCrLf & additonalInfo
    End If
End Sub

Private Sub ValidateIsInitialized()
    If Me.ParentModule Is Nothing Then
        RaiseObjectNotIntializedError
    End If
End Sub

Private Function GetErrorSource() As String
    GetErrorSource = TypeName(Me)
End Function
