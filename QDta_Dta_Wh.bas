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

Function RmvPfxzDrs(A As Drs, C$, Pfx$) As Drs
Dim Dr, ODry(), J&, I%
ODry = A.Dry
I = IxzAy(A.Fny, C)
For Each Dr In Itr(A.Dry)
    Dr(I) = RmvPfx(Dr(I), Pfx)
    ODry(J) = Dr
    J = J + 1
Next
RmvPfxzDrs = Drs(A.Fny, ODry)
End Function

Function DrswColPfx(A As Drs, C$, Pfx) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColPfx = Drs(Fny, DrywColPfx(A.Dry, Ix, Pfx))
End Function
Function LngAyzDrs(A As Drs, C$) As Long()
LngAyzDrs = IntozDrsC(EmpLngAy, A, C)
End Function
Function LngAyzDryC(Dry(), C) As Long()
LngAyzDryC = IntozDryC(EmpLngAy, Dry, C)
End Function

Function LngAyzDrswColEqSel(A As Drs, C$, V, Sel$) As Long()
LngAyzDrswColEqSel = LngAyzDrs(DrswColEqSel(A, C, V, Sel), Sel)
End Function
Function DrswColEqSel(A As Drs, C$, V, Sel$) As Drs
DrswColEqSel = SelDrs(DrswColEq(A, C, V), Sel)
End Function
Function DrswColEqSelFny(A As Drs, C$, V, SelFny$()) As Drs
DrswColEqSelFny = SelDrszFny(DrswColEq(A, C, V), SelFny)
End Function

Function DrswCCEqExlEqCol(A As Drs, CC$, V1, V2) As Drs
Dim C1$, C2$
C1 = BefSpc(CC)
C2 = AftSpc(CC)
DrswCCEqExlEqCol = DrswColEqExlEqCol(DrswColEqExlEqCol(A, C1, V1), C2, V2)
End Function
Function DrswCCCEqExlEqCol(A As Drs, CCC$, V1, V2, V3) As Drs
Dim C1$, C2$, C3$, L$
L = CCC
C1 = ShfT1(L)
C2 = ShfT1(L)
C3 = L
DrswCCCEqExlEqCol = DrswColEqExlEqCol(DrswColEqExlEqCol(DrswColEqExlEqCol(A, C1, V1), C2, V2), C3, V3)
End Function
Function DrpColzDrs(A As Drs, CC$) As Drs
Dim C$(), Dr, Ixy&(), OFny$(), ODry()
C = SyzSS(CC)
Ixy = IxyzSubAy(A.Fny, C)
OFny = MinusAy(A.Fny, C)
ODry = DrpColzDry(A.Dry, Ixy)
DrpColzDrs = Drs(OFny, ODry)
End Function
Function DrpColzDry(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrpColzDry, AyeIxy(Dr, Ixy)
Next
End Function
Function DrswColEqExlEqCol(A As Drs, C$, V) As Drs
Dim SelFny$()
SelFny = AyeEle(A.Fny, C)
DrswColEqExlEqCol = DrswColEqSelFny(A, C, V, SelFny)
End Function
Function FstDrwColEqSel(A As Drs, C$, V, Sel$) As Variant()
Stop
'FstDrwColEqSel = SelDrs(DrswColEq(A, C, V), Sel)
End Function

Function DrswColNeSel(A As Drs, C$, V, Sel$) As Drs
DrswColNeSel = SelDrs(DrswColNe(A, C, V), Sel)
End Function

Function DrswColIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
DrswColIn = Drs(A.Fny, DrywColIn(A.Dry, Ix, InVy))
End Function
Function DrswColInSel(A As Drs, C, InVy, Sel$) As Drs
DrswColInSel = SelDrs(DrswColIn(A, C, InVy), Sel)
End Function

Function DrswColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColEq = Drs(Fny, DrywColEq(A.Dry, Ix, V))
End Function

Function FstRecEqzDrs(A As Drs, C$, V, Sel$) As Drs
Dim Ix&, Ixy&(), ODry(), OFny$()
OFny = SyzSS(Sel)
Ix = IxzAy(A.Fny, C)
Ixy = IxyzSubAy(A.Fny, OFny)
ODry = FstRecEqzDry(A.Dry, Ix, V, Ixy)
FstRecEqzDrs = Drs(OFny, ODry)
End Function

Function FstDrEqzDrs(A As Drs, C$, V, Sel$) As Variant()
Dim Ix&, Ixy&(), ODry(), OFny$()
OFny = SyzSS(Sel)
Ix = IxzAy(A.Fny, C)
Ixy = IxyzSubAy(A.Fny, OFny)
FstDrEqzDrs = FstDrEqzDry(A.Dry, Ix, V, Ixy)
End Function


Function DrswColNe(A As Drs, C$, V) As Drs
Dim Dry(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DrswColNe = Drs(Fny, DrywColNe(A.Dry, Ix, V))
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

Function ValzDrswColEqSel(A As Drs, C$, V, ColNm$)
Dim Dr, Ix%, IxRet%
Ix = IxzAy(A.Fny, C)
IxRet = IxzAy(A.Fny, ColNm)
For Each Dr In Itr(A.Dry)
    If Dr(Ix) = V Then
        ValzDrswColEqSel = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with Col-A eq Value-B, so no Col-C is returened", "Col-A Value-B Col-C Drs-Fny Drs-NRec", C, V, ColNm, A.Fny, NReczDrs(A)
End Function


