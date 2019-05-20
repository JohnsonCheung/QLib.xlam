Attribute VB_Name = "QDta_Dta_Wh"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Wh."
Private Const Asm$ = "QDta"

Function DrswFldEqV(A As Drs, F, EqVal) As Drs
'DrswFldEqV = Drs(A.Fny, DryWh(A.Dry, Ixy(A.Fny, F), EqVal))
End Function

Function DrswFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DrswFFNe = Drs(Fny, DryWhCCNe(A.Dry, Ixy(Fny, F1), Ixy(Fny, F2)))
End Function
Function DrswColPfx(A As Drs, C$, Pfx) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColPfx = Drs(Fny, DrywColPfx(A.Dry, Ix, Pfx))
End Function
Function DrswColEqSel(A As Drs, C$, V, Sel$) As Drs
DrswColEqSel = SelDrs(DrswColEq(A, C, V), Sel)
End Function
Function DrswColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColEq = Drs(Fny, DrywColEq(A.Dry, Ix, V))
End Function

Function DrswColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
'Ix = Ixy(Fny, C)
DrswColGt = Drs(Fny, DrywColGt(A.Dry, Ix, V))
End Function

Function DrseRowIxy(A As Drs, RowIxy&()) As Drs
Dim ODry(), Dry()
    Dry = A.Dry
    Dim J&, I&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxy, J) Then
            PushI ODry, Dry(J)
        End If
    Next
DrseRowIxy = Drs(A.Fny, ODry)
End Function

Function DrswNotRowIxy(A As Drs, RowIxy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxy, J) Then
            Push O, Dry(J)
        End If
    Next
DrswNotRowIxy = Drs(A.Fny, O)
End Function

Function DrswRowIxy(A As Drs, RowIxy&()) As Drs
DrswRowIxy = Drs(A.Fny, CvAv(AywIxy(A.Dry, RowIxy)))
End Function

