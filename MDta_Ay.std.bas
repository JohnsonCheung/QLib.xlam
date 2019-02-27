Attribute VB_Name = "MDta_Ay"
Option Explicit

Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As DT
Dim O(), J&
For J = 0 To UB(A)
    Push O, Array(A(J))
Next
Set AyDt = DT(DtNm, FldNm, O)
End Function

Function GpCntDryzAy(A) As Variant()
If Sz(A) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AywDist(A)
For Each X In Itr(Dup)
    Cnt = AyItmCnt(A, X)
    Push O, Array(X, AyItmCnt(A, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
GpCntDryzAy = O
End Function

Function GpCntDryzAyWhDup(A) As Variant()
GpCntDryzAyWhDup = DrywCGt(GpCntDryzAy(A), 1, 1)
End Function
Sub BrwGpCntDryzAy(Ay)
Brw FmtGpCntDryzAy(Ay)
End Sub
Function FmtGpCntDryzAy(Ay) As String()
FmtGpCntDryzAy = FmtDry(GpCntDryzAy(Ay), Fmt:=eSpcSep)
End Function

Private Sub ZZ_FmtGpCntDryzAy()
Dim Ay()
Brw FmtGpCntDryzAy(Ay)
End Sub

Private Sub ZZ_CntDryzAy()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = CntDryzAy(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
Stop
'AssEqDry Act, Exp
End Sub
