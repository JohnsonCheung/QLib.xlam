Attribute VB_Name = "QIde_Cnt_SrcLin"
Option Explicit
Private Const CMod$ = "MIde_Cnt_SrcLin."
Private Const Asm$ = "QIde"


Property Get NSrcLin&()
NSrcLin = NSrcLinzPj(CurPj)
End Property

Function NSrcLinzPj&(A As VBProject)
Dim O&, C As VBComponent
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In A.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinzPj = O
End Function

