Attribute VB_Name = "QDta_Sel"
Option Explicit
Private Const CMod$ = "MDta_Sel."
Private Const Asm$ = "QDta"

Function DrySel(Dry(), Ixy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI DrySel, AywIxy(Drv, Ixy)
Next
End Function

Function SelDrs(A As Drs, FF$) As Drs
Dim Fny$(): Fny = TermAy(FF)
If IsEqAy(A.Fny, Fny) Then SelDrs = A: Exit Function
ThwNotSuperAy A.Fny, Fny
SelDrs = Drs(Fny, DrySel(A.Dry, Ixy(A.Fny, Fny)))
End Function

Private Sub Z_SelDrs()
'BrwDrs SelDrs(Vmd.MthDrs, "Mthn Mdy Ty Mdn")
'BrwDrs Vmd.MthDrs
End Sub

Function DtSel(A As Dt, FF$) As Dt
DtSel = DtzDrs(SelDrs(DrszDt(A), FF), A.DtNm)
End Function


Private Sub ZZ()
Z_SelDrs
MDta_Sel:
End Sub
