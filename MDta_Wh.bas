Attribute VB_Name = "MDta_Wh"
Option Explicit

Function DrswFldEqV(A As DRs, F, EqVal) As DRs
'Set DrswFldEqV = Drs(A.Fny, DryWh(A.Dry, IxAy(A.Fny, F), EqVal))
End Function

Function DrswFFNe(A As DRs, F1, F2) As DRs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'Set DrswFFNe = Drs(Fny, DryWhCCNe(A.Dry, IxAy(Fny, F1), IxAy(Fny, F2)))
End Function

Function DrswColEq(A As DRs, C$, V) As DRs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
'Ix = IxAy(Fny, C)
Set DrswColEq = DRs(Fny, DrywCEv(A.Dry, Ix, V))
End Function

Function DrswColGt(A As DRs, C$, V) As DRs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
'Ix = IxAy(Fny, C)
Set DrswColGt = DRs(Fny, DrywCGt(A.Dry, Ix, V))
End Function

Function DrseRowIxAy(A As DRs, RowIxAy&()) As DRs
Dim ODry(), Dry()
    Dry = A.Dry
    Dim J&, I&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxAy, J) Then
            PushI ODry, Dry(J)
        End If
    Next
Set DrseRowIxAy = DRs(A.Fny, ODry)
End Function

Function DrswNotRowIxAy(A As DRs, RowIxAy&()) As DRs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
Set DrswNotRowIxAy = DRs(A.Fny, O)
End Function


Function DrswRowIxAy(A As DRs, RowIxAy) As DRs
Set DrswRowIxAy = DRs(A.Fny, CvAv(AywIxAy(A.Dry, RowIxAy)))
End Function

