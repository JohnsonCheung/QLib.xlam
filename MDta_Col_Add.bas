Attribute VB_Name = "MDta_Col_Add"
Option Explicit

Function DrsAddColzNmVy(A As Drs, ColNm$, ColVy) As Drs
Dim Fny$(): Fny = AyAddItm(A.Fny, ColNm)
Dim AtIx&: AtIx = UB(Fny)
Dim Dry(): Dry = DryAddColzColVy(A.Dry, ColVy, AtIx)
DrsAddColzNmVy = Drs(Fny, Dry)
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

Function DrsAddColzMap(A As Drs, NewFldEqFunQuoteFmFldSsl$) As Drs
Dim NewColVy(), FmVy()
Dim I, NewFld$, Fun$, FmFld$
For Each I In SySsl(NewFldEqFunQuoteFmFldSsl)
    NewFld = Bef(I, "=")
    Fun = Bet(I, "=", "(")
    FmFld = BetBkt(I)
    FmVy = ColzDrs(A, FmFld)
    NewColVy = AyMap(FmVy, Fun)
    Stop '
Next
End Function
