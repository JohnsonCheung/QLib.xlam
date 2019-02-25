Attribute VB_Name = "MDta_ValId"
Option Explicit
Function AddColzValIdzCntzDrs(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
Dim Ix%, Fny$()
Fny = A.Fny
Ix = IxzAy(Fny, ColNm): If Ix = -1 Then Stop
    Dim X$, Y$, C$
        C = ColNmPfx & ColNm
        X = C & "Id"
        Y = C & "Cnt"
    If HasEle(Fny, X) Then Stop
    If HasEle(Fny, Y) Then Stop
    PushIAy Fny, Array(X, Y)
Set AddColzValIdzCntzDrs = Drs(Fny, AddColzValIdzCntzDry(A.Dry, Ix))
End Function

Function AddColzValIdzCntzDry(A(), ValColIx) As Variant()
Dim NCol%, Dic As Dictionary, O(), Dr, IdCnt, R&
NCol = NColDry(A)
Set Dic = IdCntDiczAy(ColzDry(A, ValColIx))
O = A
For Each Dr In Itr(A)
    ReDim Preserve Dr(NCol + 1)
    IdCnt = Dic(Dr(ValColIx))
    Dr(NCol) = IdCnt(0)
    Dr(NCol + 1) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
AddColzValIdzCntzDry = O
End Function


