Attribute VB_Name = "QIde_Cmp_is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmp_is."
Private Const Asm$ = "QIde"

Function IsCmpzMod(A As VBComponent) As Boolean
IsCmpzMod = A.Type = vbext_ct_StdModule
End Function

Function IsCmpzCls(A As VBComponent) As Boolean
IsCmpzCls = A.Type = vbext_ct_ClassModule
End Function

Function IsCmpzMd(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: IsCmpzMd = True
End Select
End Function

