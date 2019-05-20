Attribute VB_Name = "QIde_Exp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Exp."
Private Const Asm$ = "QIde"
Function ExpgPth$()
ExpgPth = PjpP & "Exporting\"
End Function

Sub ExpExpg()
Stamp "ExpExpg: Begin"
Dim Xls As Excel.Application: Set Xls = NewXls
Dim Acs As Access.Application: Set Acs = NewAcs
Dim Ffn$, I
For Each I In Itr(Ffny(ExpgPth))
    Ffn = I
    ExpPjf Ffn, Xls, Acs
Next
QuitXls Xls
QuitAcs Acs
Stamp "ExpExpg: End"
End Sub

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
Dim P$
P = Srcp(Pj)
InfLin CSub, "... Clr src pth"
    EnsPthzAllSeg P
    ClrPthFil P
InfLin CSub, "... Cpy pj to src pth"
    CpyFfnzToPth Pj.Filename, P
InfLin CSub, "... Exp src"
    ExpSrc Pj
InfLin CSub, "... Exp rf"
    ExpRf Pj
InfLin CSub, "... Exp frm"
    ExpFrm Pj
InfLin CSub, "Done"
End Sub

Private Sub ExpSrc(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    DoEvents
    C.Export SrcFfn(C)
Next
End Sub

Private Sub ExpRf(P As VBProject)
WrtAy RfSrc(P), Frf(P)
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
