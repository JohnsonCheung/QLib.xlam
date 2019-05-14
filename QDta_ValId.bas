Attribute VB_Name = "QDta_ValId"
Option Explicit
Private Const CMod$ = "MDta_ValId."
Private Const Asm$ = "QDta"
Function AddColzValIdzCntzDrs(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
Dim Ix&, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, ColNm): If Ix = -1 Then Stop
    Dim X$, Y$, C$
        C = ColNmPfx & ColNm
        X = C & "Id"
        Y = C & "Cnt"
    If HasEle(Fny, X) Then Stop
    If HasEle(Fny, Y) Then Stop
    PushIAy Fny, Array(X, Y)
AddColzValIdzCntzDrs = Drs(Fny, AddColzValIdzCntzDry(A.Dry, Ix))
End Function

Function AddColzValIdzCntzDry(Dry(), ValColIx&) As Variant()
Dim NCol%, Dic As Dictionary, O(), Dr, IdCnt, R&
NCol = NColzDry(Dry)
Set Dic = IdCntDiczAy(ColzDry(Dry, ValColIx))
O = Dry
For Each Dr In Itr(Dry)
    ReDim Preserve Dr(NCol + 1)
    IdCnt = Dic(Dr(ValColIx))
    Dr(NCol) = IdCnt(0)
    Dr(NCol + 1) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
AddColzValIdzCntzDry = O
End Function

