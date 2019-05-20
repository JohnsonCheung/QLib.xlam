Attribute VB_Name = "QIde_Export_Export"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Export."
Private Const Asm$ = "QIde"

Sub ExpMd(A As CodeModule)
A.Parent.Export SrcFfnzMd(A)
End Sub

Sub ExpRf(P As VBProject)
WrtAy RfSrc(P), Frf(P)
End Sub

Sub BrwSrcpP()
BrwPth SrcpP
End Sub

Function ExtzCmpTy$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "SrcExt: Unexpected Md_CmpTy.  Should be [Class or Module or Document]"
End Select
ExtzCmpTy = O
End Function

Function SrcFfnzMd$(A As CodeModule)
SrcFfnzMd = SrcpzP(PjzM(A)) & Mdn(A) & ExtzCmpTy(CmpTyzM(A))
End Function

Function SrcpzP$(P As VBProject)
SrcpzP = EnsPth(Pjp(P) & ".Src\" & Pjfn(P))
End Function


