Attribute VB_Name = "MIde_Export"
Option Explicit

Sub ExpMd(A As CodeModule)
A.Parent.Export SrcFfnMd(A)
End Sub

Sub ExpPjRf(A As VBProject)
WrtAy RfSrc(A), RfSrcFfn(A)
End Sub

Sub BrwPSrcPth()
BrwPth SrcPthzPj(CurPj)
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
SrcFfnMd = SrcPthzPj(PjzMd(A)) & MdNm(A) & SrcExtMd(A)
End Function

Function SrcPthzPj$(A As VBProject)
SrcPthzPj = PthEns(PjPth(A) & "Src\")
End Function


