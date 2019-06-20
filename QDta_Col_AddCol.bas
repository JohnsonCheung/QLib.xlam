Attribute VB_Name = "QDta_Col_AddCol"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Col_Add."
Private Const Asm$ = "QDta"

Function DrsAddColzNmVy(A As Drs, ColNm$, ColVy) As Drs
Dim Fny$(): Fny = AyzAddItm(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dry(): Dry = DryAddColzColVy(A.Dry, ColVy, AtIx)
DrsAddColzNmVy = Drs(Fny, Dry)
End Function
    
Function DryAddColz(Dry(), C) As Variant()
Dim O(): O = Dry
Dim ToU&
    ToU = NColzDry(Dry)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = C
    O(J) = Dr
    J = J + 1
Next
DryAddColz = O
End Function

Function DryAddColzC3(A(), C1, C2, C3) As Variant()
Dim U%, R&, Dr, O()
O = A
U = NColzDry(A) + 2
For Each Dr In Itr(A)
    ReDim Preserve Dr(U)
    Dr(U) = C3
    Dr(U - 1) = C2
    Dr(U - 2) = C1
    O(R) = Dr
    R = R + 1
Next
DryAddColzC3 = O
End Function

Function DryAddColzBy(Dry(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDry(Dry) + ByNCol - 1
Dim O()
    Dim UDry&: UDry = UB(Dry)
    O = AyReSzU(O, UDry)
    Dim J&
    For J = 0 To UDry
        O(J) = AyReSzU(Dry(J), NewU)
    Next
DryAddColzBy = O
End Function

Function DryAddColzC(Dry(), C) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim O(): O = DryAddColzBy(Dry)
    Dim UCol%: UCol = UB(Dry(0))
    Dim J&
    For J = 0 To UB(Dry)
       O(J)(UCol) = C
    Next
DryAddColzC = O
End Function

Function DryAddColzCC(Dry(), V1, V2) As Variant()
DryAddColzCC = DryAddColzAv(Dry, Av(V1, V2))
End Function

Function DryAddColzAv(Dry(), Av()) As Variant()
Dim O(): O = Dry
Dim ToU&
    ToU = NColzDry(Dry) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    PushAy Dr, Av
    O(J) = Dr
    J = J + 1
Next
DryAddColzAv = O
End Function

Function InsColzDryAv(Dry(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI InsColzDryAv, AyzAdd(Av, Dr)
Next
End Function
Function InsColzDryV3(Dry(), V1, V2, V3) As Variant()
InsColzDryV3 = InsColzDryAv(Dry, Av(V1, V2, V3))
End Function

Function InsColzDryzV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDryzV, InsEle(Dr, V, At)
Next
End Function

Function InsColzDryV4(A(), V1, V2, V3, V4) As Variant()
InsColzDryV4 = InsColzDryAv(A, Av(V1, V2, V3, V4))
End Function

Function InsColzDryV2(A(), V1, V2) As Variant()
InsColzDryV2 = InsColzDryAv(A, Av(V1, V2))
End Function



Private Function DryAddColzColVy(Dry(), ColVy, AtIx&) As Variant()
Dim Dr, J&, O(), U&
U = UB(ColVy)
If U = -1 Then Exit Function
If U <> UB(Dry) Then Thw CSub, "Row-in-Dry <> Si-ColVy", "Row-in-Dry Si-ColVy", Si(Dry), Si(ColVy)
ReDim O(U)

For Each Dr In Itr(Dry)
    If Si(Dr) > AtIx Then Thw CSub, "Some Dr in Dry has bigger size than AtIx", "DrSz AtIx", Si(Dr), AtIx
    ReDim Preserve Dr(AtIx)
    Dr(AtIx) = ColVy(J)
    O(J) = Dr
    J = J + 1
Next
DryAddColzColVy = O
End Function

Function DrsAddColzMap(A As Drs, NewFldEqFunQteFmFldSsl$) As Drs
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
