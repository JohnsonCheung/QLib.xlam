Attribute VB_Name = "QDta_F_DtaCol"
Option Explicit
Option Compare Text

Function AddColzDy(Dy(), ValToBeAddAsLasCol) As Variant()
'Ret : a new :Dy with a col of value all eq to @ValToBeAddAsLasCol at end
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = ValToBeAddAsLasCol
    O(J) = Dr
    J = J + 1
Next
AddColzDy = O
End Function

Function AddColzDyAv(Dy(), Av()) As Variant()
Dim O(): O = Dy
Dim ToU&
    ToU = NColzDy(Dy) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    PushAy Dr, Av
    O(J) = Dr
    J = J + 1
Next
AddColzDyAv = O
End Function

Function AddColzDyBy(Dy(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDy(Dy) + ByNCol - 1
Dim O()
    Dim UDy&: UDy = UB(Dy)
    O = AyReSzU(O, UDy)
    Dim J&
    For J = 0 To UDy
        O(J) = AyReSzU(Dy(J), NewU)
    Next
AddColzDyBy = O
End Function

Function AddColzDyC(Dy(), C) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim O(): O = AddColzDyBy(Dy)
    Dim UCol%: UCol = UB(Dy(0))
    Dim J&
    For J = 0 To UB(Dy)
       O(J)(UCol) = C
    Next
AddColzDyC = O
End Function

Function AddColzDyC3(A(), C1, C2, C3) As Variant()
Dim U%, R&, Dr, O()
O = A
U = NColzDy(A) + 2
For Each Dr In Itr(A)
    ReDim Preserve Dr(U)
    Dr(U) = C3
    Dr(U - 1) = C2
    Dr(U - 2) = C1
    O(R) = Dr
    R = R + 1
Next
AddColzDyC3 = O
End Function

Function AddColzDyCC(Dy(), V1, V2) As Variant()
AddColzDyCC = AddColzDyAv(Dy, Av(V1, V2))
End Function

Private Function AddColzDyColVy(Dy(), ColVy, AtIx&) As Variant()
Dim Dr, J&, O(), U&
U = UB(ColVy)
If U = -1 Then Exit Function
If U <> UB(Dy) Then Thw CSub, "Row-in-Dy <> Si-ColVy", "Row-in-Dy Si-ColVy", Si(Dy), Si(ColVy)
ReDim O(U)

For Each Dr In Itr(Dy)
    If Si(Dr) > AtIx Then Thw CSub, "Some Dr in Dy has bigger size than AtIx", "DrSz AtIx", Si(Dr), AtIx
    ReDim Preserve Dr(AtIx)
    Dr(AtIx) = ColVy(J)
    O(J) = Dr
    J = J + 1
Next
AddColzDyColVy = O
End Function

Function AddColzMap(A As Drs, NewFldEqFunQteFmFldSsl$) As Drs
Dim NewColVy(), FmVy()
Dim I, S$, NewFld$, Fun$, FmFld$
For Each I In SyzSS(NewFldEqFunQteFmFldSsl)
    S = I
    NewFld = Bef(S, "=")
    Fun = Bet(S, "=", "(")
    FmFld = BetBkt(S)
    FmVy = ColzDrs(A, FmFld)
    NewColVy = MapAy(FmVy, Fun)
    Stop '
Next
End Function

Function AddColzVy(A As Drs, ColNm$, ColVy) As Drs
Dim Fny$(): Fny = AddEle(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dy(): Dy = AddColzDyColVy(A.Dy, ColVy, AtIx)
AddColzVy = Drs(Fny, Dy)
End Function

Function CntColEq&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) = V Then O = O + 1
Next
CntColEq = O
End Function

Function CntColNe&(A As Drs, C$, V)
Dim I%: I = IxzAy(A.Fny, C)
Dim O&, Dr: For Each Dr In Itr(A.Dy)
    If Dr(I) <> V Then O = O + 1
Next
CntColNe = O
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
ColPfx = Drs(Fny, DywPfx(A.Dy, Ix, Pfx))
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

Function DrpColzDyIxy(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push DrpColzDyIxy, AeIxy(Dr, Ixy)
Next
End Function

Function FstDr(A As Drs, C$, V) As Variant()
Dim Ix&: Ix = IxzAy(A.Fny, C)
FstDr = A.Dy(Ix)
End Function

Function FstDrSel(A As Drs, C$, V, Sel$) As Variant()
FstDrSel = AwIxy(FstDr(A, C, V), IxyzAyCC(A.Fny, Sel))
End Function

Function FstRec(A As Drs, C$, V) As Drs
FstRec = Drs(A.Fny, FstDr(A, C, V))
End Function

Function HasColEq(A As Drs, C$, V) As Boolean
HasColEq = HasColEqzDy(A.Dy, IxzAy(A.Fny, C), V)
End Function

Function InsColzDyAv(Dy(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI InsColzDyAv, AddAy(Av, Dr)
Next
End Function

Function InsColzDyoV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDyoV, InsEle(Dr, V, At)
Next
End Function

Function InsColzDyV2(A(), V1, V2) As Variant()
InsColzDyV2 = InsColzDyAv(A, Av(V1, V2))
End Function

Function InsColzDyV3(Dy(), V1, V2, V3) As Variant()
InsColzDyV3 = InsColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function InsColzDyV4(A(), V1, V2, V3, V4) As Variant()
InsColzDyV4 = InsColzDyAv(A, Av(V1, V2, V3, V4))
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

Function RxyeDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it ne to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec ne @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If Not IsEqAy(Dr, Vy) Then PushI RxyeDyVy, Rix
    Rix = Rix + 1
Next
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

Function TopN(A As Drs, Optional N = 50) As Drs
TopN = Drs(A.Fny, CvAv(FstNEle(A.Dy, N)))
End Function

Function VzColEqSel(A As Drs, C$, V, ColNm$)
Dim Dr, Ix%, IxRet%
Ix = IxzAy(A.Fny, C)
IxRet = IxzAy(A.Fny, ColNm)
For Each Dr In Itr(A.Dy)
    If Dr(Ix) = V Then
        VzColEqSel = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with Col-A eq Value-B, so no Col-C is returened", "Col-A Value-B Col-C Drs-Fny Drs-NRec", C, V, ColNm, A.Fny, NReczDrs(A)
End Function

Function StrColzFst(A As Drs) As String()
StrColzFst = StrColzDy(A.Dy, 0)
End Function

Function StrColzSnd(A As Drs) As String()
StrColzSnd = StrColzDy(A.Dy, 1)
End Function

Function StrCol(A As Drs, C) As String()
StrCol = StrColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function BoolColzDrs(A As Drs, C) As Boolean()
BoolColzDrs = BoolColzDy(A.Dy, IxzAy(A.Fny, C))
End Function

Function FstColzDy(Dy()) As Variant()
FstColzDy = ColzDy(Dy, 0)
End Function

Function FstCol(A As Drs) As Variant()
FstCol = FstColzDy(A.Dy)
End Function

Function StrColzEq(A As Drs, Col$, V, ColNm$) As String()
Dim B As Drs
B = DwEqSel(A, Col, V, ColNm)
StrColzEq = StrCol(B, ColNm)
End Function

