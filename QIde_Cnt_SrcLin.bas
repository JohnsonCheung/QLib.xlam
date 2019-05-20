Attribute VB_Name = "QIde_Cnt_SrcLin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cnt_SrcLin."
Private Const Asm$ = "QIde"


Property Get NSrcLin&()
NSrcLin = NSrcLinzP(CPj)
End Property

Function NSrcLinzP&(P As VBProject)
Dim O&, C As VBComponent
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinzP = O
End Function

