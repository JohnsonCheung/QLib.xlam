Attribute VB_Name = "QDta_Dta_DtaOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Wh."
Private Const Asm$ = "QDta"

Function DwFldEqV(A As Drs, F, EqVal) As Drs
'DwFldEqV = Drs(A.Fny, DyWh(A.Dy, Ixy(A.Fny, F), EqVal))
End Function

Function DwFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DwFFNe = Drs(Fny, DyWhCCNe(A.Dy, Ixy(Fny, F1), Ixy(Fny, F2)))
End Function

Function RmvPfxzDrs(A As Drs, C$, Pfx$) As Drs
Dim Dr, ODy(), J&, I%
ODy = A.Dy
I = IxzAy(A.Fny, C)
For Each Dr In Itr(A.Dy)
    Dr(I) = RmvPfx(Dr(I), Pfx)
    ODy(J) = Dr
    J = J + 1
Next
RmvPfxzDrs = Drs(A.Fny, ODy)
End Function

Function LngAyzDrs(A As Drs, C$) As Long()
LngAyzDrs = IntozDrsC(EmpLngAy, A, C)
End Function
Function LngAyzDyC(Dy(), C) As Long()
LngAyzDyC = IntozDyC(EmpLngAy, Dy, C)
End Function

Function LngAyzColEqSel(A As Drs, C$, V, Sel$) As Long()
LngAyzColEqSel = LngAyzDrs(DwEqSel(A, C, V, Sel), Sel)
End Function
Function DwEqSel(A As Drs, C$, V, Sel$) As Drs
DwEqSel = DrszSel(DwEq(A, C, V), Sel)
End Function
Function CntColNe&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) <> V Then O = O + 1
Next
CntColNe = O
End Function
Function CntColEq&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) = V Then O = O + 1
Next
CntColEq = O
End Function
Function DwNe(A As Drs, C$, V) As Drs
DwNe = DwColNe(A, C, V)
End Function

Function ColEqSelFny(A As Drs, C$, V, SelFny$()) As Drs
ColEqSelFny = DrszSelFny(DwEq(A, C, V), SelFny)
End Function

Function DwVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be returned
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
Dim KeyDy(): KeyDy = DrszSel(D, CC).Dy
Dim Rxy&(): Rxy = RxywDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DwVap = Drs(D.Fny, ODy)
End Function

Function DeVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
DeVap = DeVy(D, CC, Vy)
End Function

Function DeVy(D As Drs, CC$, Vy) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim KeyDy(): KeyDy = DrszSel(D, CC).Dy
Dim Rxy&(): Rxy = RxyeDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DeVy = Drs(D.Fny, ODy)
End Function

Function DwCCEqExlEqCol(A As Drs, CC$, V1, V2) As Drs
Dim C1$, C2$
C1 = BefSpc(CC)
C2 = AftSpc(CC)
DwCCEqExlEqCol = ColEqExlEqCol(ColEqExlEqCol(A, C1, V1), C2, V2)
End Function

Function DwCCEq(A As Drs, CC$, V1, V2) As Drs
Dim C1$, C2$
C1 = BefSpc(CC)
C2 = AftSpc(CC)
DwCCEq = DwEq(DwEq(A, C1, V1), C2, V2)
End Function

Function DwCCCEqExlEqCol(A As Drs, CCC$, V1, V2, V3) As Drs
Dim C1$, C2$, C3$, L$
L = CCC
C1 = ShfT1(L)
C2 = ShfT1(L)
C3 = L
DwCCCEqExlEqCol = ColEqExlEqCol(ColEqExlEqCol(ColEqExlEqCol(A, C1, V1), C2, V2), C3, V3)
End Function
Function DrpCol(A As Drs, CC$) As Drs
Dim C$(), Dr, Ixy&(), OFny$(), ODy()
C = SyzSS(CC)
Ixy = IxyzSubAy(A.Fny, C)
OFny = MinusAy(A.Fny, C)
ODy = DrpColzDy(A.Dy, Ixy)
DrpCol = Drs(OFny, ODy)
End Function

Function DrpColzDy(Dy(), Ixy&()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DrpColzDy, AeIxy(Dr, Ixy)
Next
End Function
Function ColEqExlEqCol(A As Drs, C$, V) As Drs
Dim SelFny$()
SelFny = AeEle(A.Fny, C)
ColEqExlEqCol = ColEqSelFny(A, C, V, SelFny)
End Function
Function DwNeSel(A As Drs, C$, V, Sel$) As Drs
DwNeSel = DrszSel(DwNe(A, C, V), Sel)
End Function

Function DeIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    If Not HasEle(InVy, Dr(Ix)) Then
        PushI Dy, Dr
    End If
Next
DeIn = Drs(A.Fny, Dy)
End Function

Function DwIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
DwIn = Drs(A.Fny, DywIn(A.Dy, Ix, InVy))
End Function
Function ColInSel(A As Drs, C, InVy, Sel$) As Drs
ColInSel = DrszSel(DwIn(A, C, InVy), Sel)
End Function

Function DwEq(A As Drs, C$, V) As Drs
DwEq = DwColEq(A, C, V)
End Function

Function DwDup(D As Drs, FF$) As Drs
DwDup = DwRxy(D, RxyzDup(D, FF))
End Function

Function DwDupC(A As Drs, C$) As Drs
Dim Dup(): Dup = AwDup(ColzDrs(A, C))
DwDupC = DwIn(A, C, Dup)
End Function

Function DwEqExl(A As Drs, C$, V) As Drs
DwEqExl = DrpCol(DwColEq(A, C, V), C)
End Function
Function TopN(A As Drs, Optional N = 50) As Drs
TopN = Drs(A.Fny, CvAv(FstNEle(A.Dy, N)))
End Function
Function ColNoSng(A As Drs, C$) As Drs
'Fm  A : has a column-C
'Ret   : sam stru as A and som row removed.  rmv row are its col C value is Single. @@
Dim Col(): Col = ColzDrs(A, C)
Dim Sng(): Sng = AwSng(Col)
ColNoSng = DeIn(A, C, Sng)
End Function

Function ColPfx(A As Drs, C$, Pfx$) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
ColPfx = Drs(Fny, DywColPfx(A.Dy, Ix, Pfx))
End Function
Function Dw2Patn(A As Drs, TwoC$, Patn1$, Patn2$) As Drs
Dim C1$, C2$: AsgBrkSpc TwoC, C1, C2
Dw2Patn = DwPatn(DwPatn(A, C1, Patn1), C2, Patn2)
End Function
Function DwPatn(A As Drs, C$, Patn$) As Drs
Dim I%: I = IxzAy(A.Fny, C, ThwEr:=EiThwEr)
Dim Re As RegExp: Set Re = RegExp(Patn)
Dim Dy(), Dr: For Each Dr In Itr(A.Dy)
    If Re.Test(Dr(I)) Then PushI Dy, Dr
Next
DwPatn = Drs(A.Fny, Dy)
End Function
Function DTopN(A As Drs, Optional NTop& = 50) As Drs
Dim Dy(): Dy = FstNEle(A.Dy, NTop)
DTopN = Drs(A.Fny, Dy)
End Function
Function DwColEq(A As Drs, C$, V) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwColEq = Drs(Fny, DywColEq(A.Dy, Ix, V))
End Function
Function HasColEq(A As Drs, C$, V) As Boolean
HasColEq = HasColEqzDy(A.Dy, IxzAy(A.Fny, C), V)
End Function

Function FstRecEqzDrs(A As Drs, C$, V, Sel$) As Drs
Dim Ix&, Ixy&(), ODy(), OFny$()
OFny = SyzSS(Sel)
Ix = IxzAy(A.Fny, C)
Ixy = IxyzSubAy(A.Fny, OFny)
ODy = FstRecEqzDy(A.Dy, Ix, V, Ixy)
FstRecEqzDrs = Drs(OFny, ODy)
End Function

Function FstDrEqzDrs(A As Drs, C$, V, Sel$) As Variant()
Dim Ix&, Ixy&(), ODy(), OFny$()
OFny = SyzSS(Sel)
Ix = IxzAy(A.Fny, C)
Ixy = IxyzSubAy(A.Fny, OFny)
FstDrEqzDrs = FstDrEqzDy(A.Dy, Ix, V, Ixy)
End Function

Function DwColNe(A As Drs, C$, V) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DwColNe = Drs(Fny, DywColNe(A.Dy, Ix, V))
End Function

Function DwColGt(A As Drs, C$, V) As Drs
Dim Dy(), Ix%, Fny$()
Fny = A.Fny
'Ix = Ixy(Fny, C)
DwColGt = Drs(Fny, DywColGt(A.Dy, Ix, V))
End Function

Function DeRxy(A As Drs, Rxy&()) As Drs
Dim ODy(), Dy()
    Dy = A.Dy
    Dim J&, I&
    For J = 0 To UB(Dy)
        If Not HasEle(Rxy, J) Then
            PushI ODy, Dy(J)
        End If
    Next
DeRxy = Drs(A.Fny, ODy)
End Function

Function DwNotRxy(A As Drs, Rxy&()) As Drs
Dim O(), Dy()
    Dy = A.Dy
    Dim J&
    For J = 0 To UB(Dy)
        If Not HasEle(Rxy, J) Then
            Push O, Dy(J)
        End If
    Next
DwNotRxy = Drs(A.Fny, O)
End Function

Function DwRxy(A As Drs, Rxy&()) As Drs
Dim Dy(): Dy = AwIxy(A.Dy, Rxy)
DwRxy = Drs(A.Fny, Dy)
End Function

Function ValzColEqSel(A As Drs, C$, V, ColNm$)
Dim Dr, Ix%, IxRet%
Ix = IxzAy(A.Fny, C)
IxRet = IxzAy(A.Fny, ColNm)
For Each Dr In Itr(A.Dy)
    If Dr(Ix) = V Then
        ValzColEqSel = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with Col-A eq Value-B, so no Col-C is returened", "Col-A Value-B Col-C Drs-Fny Drs-NRec", C, V, ColNm, A.Fny, NReczDrs(A)
End Function

Function RxywDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it eq to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec eq @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If IsEqAy(Dr, Vy) Then PushI RxywDyVy, Rix
    Rix = Rix + 1
Next
End Function
Function RxyeDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it ne to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec ne @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If Not IsEqAy(Dr, Vy) Then PushI RxyeDyVy, Rix
    Rix = Rix + 1
Next
End Function

