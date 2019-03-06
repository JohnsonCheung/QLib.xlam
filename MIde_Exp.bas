Attribute VB_Name = "MIde_Exp"
Option Explicit
Function ExpgPth$()
ExpgPth = PthPj & "Exporting\"
End Function

Sub ExpExpg()
Stamp "ExpExpg: Begin"
Dim Xls As Excel.Application: Set Xls = NewXls
Dim Acs As Access.Application: Set Acs = NewAcs
Dim Ffn
For Each Ffn In Itr(FfnAy(ExpgPth))
    ExpPjf Ffn, Xls, Acs
Next
XlsQuit Xls
AcsQuit Acs
Stamp "ExpExpg: End"
End Sub

Sub ExpPjf(Pjf, Optional Xls As Excel.Application, Optional Acs As Access.Application)
Stamp "ExpPj: Begin"
Stamp "ExpPj: Pjf " & Pjf
Select Case True
Case IsFxa(Pjf): ExpFx Pjf, Xls
Case IsFba(Pjf): ExpFb Pjf, Acs
End Select
Stamp "ExpPj: End"
End Sub

Sub Z1()
ExpExpg
End Sub

Sub ExpFb(Fb, Optional Acs As Access.Application)
CpyFilzToPth Fb, PthEns(SrcPthzPjf(Fb))
Dim A As Access.Application: Set A = DftAcs(Acs)
OpnFb A, Fb
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
ExpzPj Pj
If IsNothing(Acs) Then AcsQuit A
End Sub

Sub ExpFx(Fx, Optional Xls As Excel.Application)
CpyFilzToPth Fx, PthEns(SrcPthzPjf(Fx))
Dim A As Excel.Application: Set A = DftXls(Xls)
A.Workbooks.Open Fx
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
ExpzPj Pj
If IsNothing(Xls) Then XlsQuit A
End Sub

Sub ExpPj()
ExpzPj CurPj
End Sub

Sub ExpzPj(Pj As VBProject)
Dim P$: P = SrcPthEns(Pj)
ClrPthFil P
ExpSrc Pj
ExpRf Pj
ExpFrm Pj
ZipPth P
End Sub

Private Sub ExpSrc(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    DoEvents
    C.Export SrcFfn(C)
Next
End Sub

Private Sub ExpRf(A As VBProject)
WrtAy RfSrc(A), RfSrcFfn(A)
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

