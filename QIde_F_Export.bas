Attribute VB_Name = "QIde_F_Export"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Export."
Private Const Asm$ = "QIde"

Sub ExpMd(M As CodeModule)
M.Parent.Export SrcFfnzMd(M)
End Sub

Private Sub ExpRf(P As VBProject)
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

Function SrcFfnzMd$(M As CodeModule)
SrcFfnzMd = SrcpzP(PjzM(M)) & Mdn(M) & ExtzCmpTy(CmpTyzM(M))
End Function

Function SrcpzP$(P As VBProject)
SrcpzP = EnsPth(Pjp(P) & ".Src\" & Pjfn(P))
End Function



Sub ExpPjf(Pjf, Optional Xls As Excel.Application, Optional Acs As Access.Application)
Stamp "ExpPj: Begin"
Stamp "ExpPj: Pjf " & Pjf
Select Case True
Case IsFxa(Pjf): ExpFxa Pjf
Case IsFba(Pjf): ExpFba Pjf, Acs
End Select
Stamp "ExpPj: End"
End Sub

Sub ExpFba(Fba, Optional Acs As Access.Application)
CpyFfnzToPth Fba, EnsPth(SrcpzPjf(Fba))
Dim A As Access.Application: Set A = DftAcs(Acs)
OpnFb A, Fba
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
ExpPj Pj
QuitAcs A
End Sub

Sub ExpFxa(Fxa)
ExpPj PjzFxa(Fxa)
End Sub

Sub ExpP()
ExpPj CPj
End Sub

Sub ExpPj(Pj As VBProject)
Dim P$: P = Srcp(Pj)
InfLin CSub, "... Clr src pth":       EnsPthzAllSeg P
                                      ClrPthFil P
InfLin CSub, "... Cpy pj to src pth": CpyFfnzToPth Pj.Filename, P
InfLin CSub, "... Exp src":           ExpSrc Pj
InfLin CSub, "... Exp rf":            ExpRf Pj
InfLin CSub, "... Exp frm":           ExpFrm Pj
InfLin CSub, "Done"
End Sub

Private Sub ExpSrc(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    DoEvents
    C.Export SrcFfn(C)
Next
End Sub

Private Sub ExpFrm(P As VBProject)
If Not IsFbaPj(P) Then Exit Sub
Stop
End Sub

Private Sub ExpFrmzAcs(A As Access.Application, ToPth$)
Dim F As AccessObject
For Each F In A.CurrentProject.AllForms
    A.SaveAsText acForm, F.Name, ToPth & F.Name & ".frm.txt"
Next
End Sub

Private Sub Z()
QIde_F_Export.BrwSrcpP

End Sub
