Attribute VB_Name = "QIde_Cmp_is"
Option Explicit
Private Const CMod$ = "MIde_Cmp_is."
Private Const Asm$ = "QIde"

Function IsModCmp(A As VBComponent) As Boolean
IsModCmp = A.Type = vbext_ct_StdModule
End Function

Function IsClsCmp(A As VBComponent) As Boolean
IsClsCmp = A.Type = vbext_ct_ClassModule
End Function


Function IsMd(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: IsMd = True
End Select
End Function

