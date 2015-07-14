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
'
' Delegate
' ========
'
' Black Magic...

''
' CopyMemory: We plan to copy the actual bits of the compiled function
' to the compiled bits of the private `Run` function.
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)

' http://www.codeproject.com/Articles/19239/Compilers-Demystified-Function-Pointers-in-Visual
'
' Private Members
' ---------------
Private Const OFFSET_TO_VFT = &H1C
Private Const VFT_ELEM_SIZE As Integer = 4

Private pProcAddress As Long ' address to copy to pRunVFT
Private pBlankProcVFTLoc As Long ' location _of_ `BlankProc` on VFTable
Private pBlankProcAddress As Long ' address _at_ pBlankProcVFTLoc on VFTable
'
' BLANK PROC
' ----------
'
' This procedure will be transfigured with black magic
' it needs to be at the top ALWAYS!!!!!
' casue black magic...
Public Sub BlankProc()
    Err.Raise 1
End Sub
'
' Constructors
' ------------
'
Private Sub Class_Initialize()
    
    ' get VFT table
    Dim vft As Long ' location of vftable
    CopyMemory vft, ByVal ObjPtr(Me), VFT_ELEM_SIZE
    
    ' calculate `BlankProc` location on VFT table
    pBlankProcVFTLoc = vft + OFFSET_TO_VFT
    
    ' store Address of `BlankProc` from VFT table
    CopyMemory pBlankProcAddress, ByVal pBlankProcVFTLoc, VFT_ELEM_SIZE
    
End Sub
Public Function Create(ByVal procAddr As Long) As Delegate

    Dim result As New Delegate
    result.ProcedureAddress = procAddr
    Set Create = result
    
End Function
Friend Property Let ProcedureAddress(ByVal procAddr As Long)
    
    pProcAddress = procAddr
    
End Property
Public Sub Run()
    
    ' > One more important thing to remember when modifying vftables,
    ' > is that they are shared by all instances of that class.
    '
    ' ergo we must modify the global definition before we run and
    ' return it to normal afterwards.
    CopyMemory ByVal pBlankProcVFTLoc, pProcAddress, VFT_ELEM_SIZE
    BlankProc
    CopyMemory ByVal pBlankProcVFTLoc, pBlankProcAddress, VFT_ELEM_SIZE
    ' This is about as un-threadsafe as one can be...
    
End Sub




