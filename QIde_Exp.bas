Attribute VB_Name = "QIde_Exp"
Option Explicit
Private Const CMod$ = "MIde_Exp."
Private Const Asm$ = "QIde"
Function ExpgPth$()
ExpgPth = PthPj & "Exporting\"
End Function

Sub ExpExpg()
Stamp "ExpExpg: Begin"
Dim Xls As Excel.Application: Set Xls = NewXls
Dim Acs As Access.Application: Set Acs = NewAcs
Dim Ffn$, I
For Each I In Itr(FfnSy(ExpgPth))
    Ffn = I
    ExpPjf Ffn, Xls, Acs
Next
QuitXls Xls
QuitAcs Acs
Stamp "ExpExpg: End"
End Sub

Sub ExpPjf(Pjf$, Optional Xls As Excel.Application, Optional Acs As Access.Application)
Stamp "ExpPj: Begin"
Stamp "ExpPj: Pjf " & Pjf
Select Case True
Case IsFxa(Pjf): ExpFxa Pjf, Xls
Case IsFba(Pjf): ExpFb Pjf, Acs
End Select
Stamp "ExpPj: End"
End Sub

Sub ExpFb(Fb$, Optional Acs As Access.Application)
CpyFfnzToPth Fb, EnsPth(SrcpzPjf(Fb$))
Dim A As Access.Application: Set A = DftAcs(Acs)
OpnFb A, Fb
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
PjExp Pj
If IsNothing(Acs) Then QuitAcs A
End Sub

Sub ExpFxa(Fxa$)
ExpPj PjzFxa(Fxa)
End Sub

Sub ExpP()
ExpPj CurPj
End Sub

Sub ExpPj(Pj As VBProject)
Dim P$
P = Srcp(Pj)
EnsPthzAllSeg P
ClrPthFil P
CpyFfnzToPth Pj.Filename, P
ExpSrc Pj
ExpRf Pj
ExpFrm Pj
End Sub

Private Sub ExpSrc(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    DoEvents
    C.Export SrcFfn(C)
Next
End Sub

Private Sub ExpRf(A As VBProject)
WrtAy RfSrc(A), Frf(A)
End Sub

Private Sub ExpFrm(A As VBProject)
If Not IsFbaPj(A) Then Exit Sub
Stop
End Sub

Private Sub ExpFrmzAcs(A As Access.Application, ToPth$)
Dim F As AccessObject
For Each F In A.CurrentProject.AllForms
    A.SaveAsText acForm, F.Name, ToPth & F.Name & ".frm.txt"
Next
End Sub
