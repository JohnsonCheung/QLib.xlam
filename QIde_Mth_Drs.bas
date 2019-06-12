Attribute VB_Name = "QIde_Mth_Drs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Drs."
Private Const Asm$ = "QIde"

Function DMthCzFxa(Fxa$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
DMthCzFxa = DMthCzP(PjzFxa(Fxa))
If IsNothing(Xls) Then QuitXls Xls
End Function

Function DMthCzM(M As CodeModule) As Drs
'DMthCzMd = Drs(MthFny, MthDryzMd(M))
Dim P$, T$, N$
P = PjnzM(M)
T = ShtCmpTyzMd(M)
N = Mdn(M)
'DMthCzM = DryInsColzV3(MthDryzS(Src(M)), P, T, N)
End Function

Function DMthCzP(P As VBProject) As Drs
Dim O As Drs
'O = Drs(MthFny, MthDryzP(P))
'O = AddColzValIdzCntzDrs(O, "Lines", "Pj")
'O = AddColzValIdzCntzDrs(O, "Nm", "PjMth")
'MthDrszP = O
End Function

Function DMthCzPjf(Pjf) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
DMthCzPjf = DrsAddCol(DMthCzP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function DMthCzPjfy(Pjfy$()) As Drs
Dim F
For Each F In Pjfy
    ApdDrs DMthCzPjfy, DMthCzPjf(F)
Next
End Function

Function DMthCzV(V As Vbe) As Drs
Dim P As VBProject: For Each P In V.VBProjects
    Dim A As Drs: A = DMthCzP(P)
    Dim O As Drs: O = AddDrs(O, A)
Next
DMthCzV = O
End Function

Function Dr_MthLin(MthLin) As Variant()
'If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
Dr_MthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function
