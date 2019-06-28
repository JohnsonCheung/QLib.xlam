Attribute VB_Name = "QDta_B_DtaOp"
Option Explicit
Option Compare Text


Function DrpColzDy(Dy(), C&) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI DrpColzDy, AeEleAt(Dr, C)
Next
End Function

Function DrpColzDyIxy(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push DrpColzDyIxy, AeIxy(Dr, Ixy)
Next
End Function

Function AddColByColVy(A As Drs, ColNm$, ColVy) As Drs
Dim Fny$(): Fny = AddAyItm(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dy(): Dy = AddColzDyColVy(A.Dy, ColVy, AtIx)
AddColByColVy = Drs(Fny, Dy)
End Function
    
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

Function AddColzDyCC(Dy(), V1, V2) As Variant()
AddColzDyCC = AddColzDyAv(Dy, Av(V1, V2))
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

Function InsColzDyAv(Dy(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI InsColzDyAv, AddAy(Av, Dr)
Next
End Function
Function InsColzDyV3(Dy(), V1, V2, V3) As Variant()
InsColzDyV3 = InsColzDyAv(Dy, Av(V1, V2, V3))
End Function

Function InsColzDyoV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDyoV, InsEle(Dr, V, At)
Next
End Function

Function InsColzDyV4(A(), V1, V2, V3, V4) As Variant()
InsColzDyV4 = InsColzDyAv(A, Av(V1, V2, V3, V4))
End Function

Function InsColzDyV2(A(), V1, V2) As Variant()
InsColzDyV2 = InsColzDyAv(A, Av(V1, V2))
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

