Attribute VB_Name = "MDta_Wh"
Option Explicit

Function DrswFldEqV(A As Drs, F, EqVal) As Drs
'DrswFldEqV = Drs(A.Fny, DryWh(A.Dry, IxAy(A.Fny, F), EqVal))
End Function

Function DrswFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DrswFFNe = Drs(Fny, DryWhCCNe(A.Dry, IxAy(Fny, F1), IxAy(Fny, F2)))
End Function

Function DrswColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColEq = Drs(Fny, DrywCEv(A.Dry, Ix, V))
End Function

Function DrswColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
'Ix = IxAy(Fny, C)
DrswColGt = Drs(Fny, DrywCGt(A.Dry, Ix, V))
End Function

Function DrseRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim ODry(), Dry()
    Dry = A.Dry
    Dim J&, I&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxAy, J) Then
            PushI ODry, Dry(J)
        End If
    Next
DrseRowIxAy = Drs(A.Fny, ODry)
End Function

Function DrswNotRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not HasEle(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
DrswNotRowIxAy = Drs(A.Fny, O)
End Function

Function DrswRowIxAy(A As Drs, RowIxAy&()) As Drs
DrswRowIxAy = Drs(A.Fny, CvAv(AywIxAy(A.Dry, RowIxAy)))
End Function

