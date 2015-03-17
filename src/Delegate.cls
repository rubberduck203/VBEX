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

''
' CopyMemory: We plan to copy the actual bits of the compiled function
' to the compiled bits of the private `Run` function.
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)

' http://www.codeproject.com/Articles/19239/Compilers-Demystified-Function-Pointers-in-Visual
Private Const OFFSET_TO_VFT = &H1C
Private Const VFT_ELEM_SIZE As Integer = 4


Private pVFT As Long ' location of
Private pProcAddress As Long ' address to copy to pRunVFT
Private pRunVFT As Long ' address of `Run` on VFT table
'Private pBlankProcAddress As Long ' address at pRunVFT on VFT table

Private Sub Class_Initialize()
    
    ' get VFT table
    CopyMemory pVFT, ByVal ObjPtr(Me), VFT_ELEM_SIZE
    
    ' calculate `Run` location on VFT table
    pRunVFT = pVFT + OFFSET_TO_VFT + VFT_ELEM_SIZE
    
    ' store Address of `Run` from VFT table
    ' CopyMemory pBlankProcAddress, ByVal pRunVFT, VFT_ELEM_SIZE

End Sub
' TODO: determine func datatype
Public Function Create(ByVal procAddr As Long) As Delegate

    Dim result As New Delegate
    result.ProcedureAddress = procAddr
    Set Create = result
    
End Function
Friend Property Let ProcedureAddress(ByVal procAddr As Long)
    
    pProcAddress = procAddr
    
End Property
Public Sub BlankProc()
    Err.Raise 1
End Sub
Public Sub Run()
    
    CopyMemory ByVal pRunVFT, pProcAddress, VFT_ELEM_SIZE
    BlankProc
    'CopyMemory ByVal pRunVFT, pBlankProcAddress, VFT_ELEM_SIZE
    
End Sub
