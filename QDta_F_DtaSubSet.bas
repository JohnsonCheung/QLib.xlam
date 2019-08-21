Attribute VB_Name = "QDta_F_DtaSubSet"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Wh."
Private Const Asm$ = "QDta"

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

Function DeVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
DeVap = DeVy(D, CC, Vy)
End Function

Function DeVy(D As Drs, CC$, Vy) As Drs
'Fm D  : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vy : #Val-Ay-of-Param ! to select what rec in @D to be selected
'Ret   : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxyeDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DeVy = Drs(D.Fny, ODy)
End Function
Function Dw2Eq(D As Drs, C2$, V1, V2) As Drs
Dim A$, B$: AsgTRst C2, A, B
Dw2Eq = DwEq(DwEq(D, A, V1), B, V2)
End Function

Function Dw2EqE(D As Drs, C2$, V1, V2) As Drs
Dw2EqE = DrpCol(Dw2Eq(D, C2, V1, V2), C2)
End Function

Function Dw2Patn(A As Drs, TwoC$, Patn1$, Patn2$) As Drs
Dim C1$, C2$: AsgBrkSpc TwoC, C1, C2
Dw2Patn = DwPatn(DwPatn(A, C1, Patn1), C2, Patn2)
End Function

Function Dw3Eq(D As Drs, C3$, V1, V2, V3) As Drs
Dim A$, B$, C$: AsgTTRst C3, A, B, C
Dw3Eq = DwEq(DwEq(DwEq(D, A, V1), B, V2), C, V3)
End Function

Function Dw3EqE(D As Drs, C3$, V1, V2, V3) As Drs
Dw3EqE = DrpCol(Dw3Eq(D, C3, V1, V2, V3), C3)
End Function

Function DwColGt(A As Drs, C$, V) As Drs
Dim Dy(), Ix%, Fny$()
Fny = A.Fny
'Ix = Ixy(Fny, C)
DwColGt = Drs(Fny, DywColGt(A.Dy, Ix, V))
End Function

Function DwColNe(A As Drs, C$, V) As Drs
Dim Dy(), Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C)
DwColNe = Drs(Fny, DywColNe(A.Dy, Ix, V))
End Function

Function DwDup(D As Drs, FF$) As Drs
DwDup = DwRxy(D, RxyzDup(D, FF))
End Function

Function DwDupC(A As Drs, C$) As Drs
Dim Dup(): Dup = AwDup(ColzDrs(A, C))
DwDupC = DwIn(A, C, Dup)
End Function

Function DwBlnk(A As Drs, C$) As Drs
DwBlnk = DwEq(A, C, "")
End Function

Function DwNBlnk(A As Drs, C$) As Drs
DwNBlnk = DwNe(A, C, "")
End Function

Function DwEq(A As Drs, C$, V) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwEq = Drs(Fny, DywEq(A.Dy, Ix, V))
End Function

Function DwEqStr(A As Drs, C$, Str$) As Drs
If Str = "" Then DwEqStr = A: Exit Function
DwEqStr = DwEq(A, C, Str)
End Function

Function DwSubStr(A As Drs, C$, SubStr) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwSubStr = Drs(Fny, DywSubStr(A.Dy, Ix, SubStr))
End Function

Function DwLik(A As Drs, C$, Lik) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, C, ThwEr:=EiThwEr)
DwLik = Drs(Fny, DywLik(A.Dy, Ix, Lik))
End Function

Function DwEqE(A As Drs, C$, V) As Drs
Dim SelFny$()
SelFny = AeEle(A.Fny, C)
DwEqE = DwEqFny(A, C, V, SelFny)
End Function

Function DwEqExl(A As Drs, C$, V) As Drs
DwEqExl = DrpCol(DwEq(A, C, V), C)
End Function

Function DwEqSel(A As Drs, C$, V, Sel$) As Drs
DwEqSel = SelDrs(DwEq(A, C, V), Sel)
End Function

Function DwFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
Dim Fny$()
Fny = A.Fny
'DwFFNe = Drs(Fny, DyWhCCNe(A.Dy, Ixy(Fny, F1), Ixy(Fny, F2)))
End Function

Function DwFldEqV(A As Drs, F, EqVal) As Drs
'DwFldEqV = Drs(A.Fny, DyWh(A.Dy, Ixy(A.Fny, F), EqVal))
End Function

Function DwIn(A As Drs, C, InVy) As Drs
Dim Ix&: Ix = IxzAy(A.Fny, C)
DwIn = Drs(A.Fny, DywIn(A.Dy, Ix, InVy))
End Function

Function DwNe(A As Drs, C$, V) As Drs
DwNe = DwColNe(A, C, V)
End Function

Function DwNeSel(A As Drs, C$, V, Sel$) As Drs
DwNeSel = SelDrs(DwNe(A, C, V), Sel)
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

Function DwPatn(A As Drs, C$, Patn$, Optional Patn1$, Optional Patn2$) As Drs
If Patn = "" Then DwPatn = A: Exit Function
Dim I%: I = IxzAy(A.Fny, C, ThwEr:=EiThwEr)
Dim P As IPred: Set P = PredHasPatn(Patn, Patn1, Patn2)
Dim Dy(), Dr: For Each Dr In Itr(A.Dy)
    If P.Pred(Dr(I)) Then PushI Dy, Dr
Next
DwPatn = Drs(A.Fny, Dy)
End Function

Function DwPfx(D As Drs, C$, Pfx) As Drs
DwPfx = Drs(D.Fny, DywPfx(D.Dy, IxzAy(D.Fny, C), Pfx))
End Function

Function DwRxy(A As Drs, Rxy&()) As Drs
Dim Dy(): Dy = AwIxy(A.Dy, Rxy)
DwRxy = Drs(A.Fny, Dy)
End Function

Function DwTop(A As Drs, Optional NTop& = 50) As Drs
DwTop = Drs(A.Fny, CvAv(FstNEle(A.Dy, NTop)))
End Function

Function DwVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be returned
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
Dim KeyDy(): KeyDy = SelDrs(D, CC).Dy
Dim Rxy&(): Rxy = RxywDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DwVap = Drs(D.Fny, ODy)
End Function

Function DwEqFny(A As Drs, C$, V, SelFny$()) As Drs
DwEqFny = SelDrsFny(DwEq(A, C, V), SelFny)
End Function

Function DwInSel(A As Drs, C, InVy, Sel$) As Drs
DwInSel = SelDrs(DwIn(A, C, InVy), Sel)
End Function


