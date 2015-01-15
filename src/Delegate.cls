VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Delegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
Implements IApplicable
'
'

''
' CopyMemory: We plan to copy the actual bits of the compiled function 
' to the compiled bits of the private `Run` function.
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

''
' FillMemory: For safety we will fill the `Run` function with zero when
' the instance is destroyed.
Private Declare PtrSafe Sub FillMemory Lib "kernel32" _
    Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)


' http://www.codeproject.com/Articles/19239/Compilers-Demystified-Function-Pointers-in-Visual 
Private Const OFFSET_TO_VFT = &H1C
Private Const FUNC_ADDRESS_SIZE As Long = ???
Private Const VFT_SIZE As Integer = 4
Private Const MEMORY_FILLER As Byte = 0
Private pFuncPtr As Long


Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()

    FillMemory pFuncPtr, FUNC_ADDRESS_LENGTH, MEMORY_FILLER

End Sub

' TODO: determine func datatype
Public Function Create(ByVal func As Variant) As Delegate
    
End Function

Public Function Copy(ByVal other As Delegate) As Delegate

End Function

Private Function Run(???) As Variant
End Function

Public Function Apply(ByRef args() As Variant) As Variant

  cast.Assign Apply, Run(???)

End Function

Public Function Exec(ParamArray args() As Variant) As Variant

    cast.Assign Exec, Apply(cast.CArr(args))

End Function
'
' Private Properties
' ------------------
'
Private Property Get VFTAddress() As Long

    CopyMemory VFTAddress, (ObjPtr(Me)), VFT_SIZE

End Property
'
'
' ### Interfaces
'
Public Function IApplicable_Apply(ByRef args() As Variant) As Variant

    IApplicable_Apply = Apply(args)

End Function

Public Function IApplicable_Exec(ParamArray args() As Variant) As Variant

    IApplicable_Exec = Apply(cast.CArr(args))

End Function
