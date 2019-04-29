Attribute VB_Name = "MIde_Export"
Option Explicit

Sub ExpMd(A As CodeModule)
A.Parent.Export SrcFfnMd(A)
End Sub

Sub ExpPjRf(A As VBProject)
WrtAy RfSrczPj(A), RfSrcFfn(A)
End Sub

Sub BrwPSrcp()
BrwPth SrcpzPj(CurPj)
End Sub

Function SrcExtMd$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "SrcExt: Unexpected Md_CmpTy.  Should be [Class or Module or Document]"
End Select
SrcExtMd = O
End Function

Function SrcFfnMd$(A As CodeModule)
SrcFfnMd = SrcpzPj(PjzMd(A)) & MdNm(A) & SrcExtMd(A)
End Function

Function SrcpzPj$(A As VBProject)
SrcpzPj = EnsPth(PjPth(A) & ".Src\" & Pjfn(A))
End Function


