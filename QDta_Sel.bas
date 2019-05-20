Attribute VB_Name = "QDta_Sel"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Sel."
Private Const Asm$ = "QDta"

Function SelDry(Dry(), Ixy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI SelDry, AywIxy(Drv, Ixy)
Next
End Function
Function ExpandFF(FF$, Fny$()) As String() '
ExpandFF = ExpandLikAy(TermAy(FF), Fny)
End Function
Function ExpandLikAy(LikAy$(), Ay$()) As String() 'Put each expanded-ele in likAy to return a return ay. _
Expanded-ele means either the ele itself if there is no ele in Ay is like the `ele` _
                   or     the lik elements in Ay with the given `ele`
Dim Lik
For Each Lik In LikAy
    Dim A$()
    A = AywLik(Ay, Lik)
    If Si(A) = 0 Then
        PushI ExpandLikAy, Lik
    Else
        PushIAy ExpandLikAy, A
    End If
Next
End Function
Function SelDrs(A As Drs, FF$) As Drs
Dim Fny$(): Fny = ExpandFF(FF, A.Fny)
ThwNotSuperAy A.Fny, Fny
SelDrs = SelDrsAlwEmpzFny(A, Fny)
End Function

Function SelDrsAlwEmpzFny(A As Drs, Fny$()) As Drs
If IsEqAy(A.Fny, Fny) Then SelDrsAlwEmpzFny = A: Exit Function
SelDrsAlwEmpzFny = Drs(Fny, SelDry(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function SelDrsAlwEmp(A As Drs, FF$) As Drs
SelDrsAlwEmp = SelDrsAlwEmpzFny(A, TermAy(FF))
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
