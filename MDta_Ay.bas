Attribute VB_Name = "MDta_Ay"
Option Explicit

Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim O(), J&
For J = 0 To UB(A)
    Push O, Array(A(J))
Next
Set AyDt = Dt(DtNm, FldNm, O)
End Function

Function DryGpCntzAy(A) As Variant()
If Si(A) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AywDist(A)
For Each X In Itr(Dup)
    Cnt = AyItmCnt(A, X)
    Push O, Array(X, AyItmCnt(A, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
DryGpCntzAy = O
End Function

Function DryGpCntzAyWhDup(A) As Variant()
DryGpCntzAyWhDup = DrywCGt(DryGpCntzAy(A), 1, 1)
End Function
Sub BrwDryGpCntzAy(Ay)
Brw FmtDryGpCntzAy(Ay)
End Sub

Function FmtDryGpCntzAy(Ay) As String()
FmtDryGpCntzAy = FmtDryAsSpcSep(DryGpCntzAy(Ay))
End Function

Private Sub ZZ_FmtDryGpCntzAy()
Dim Ay()
Brw FmtDryGpCntzAy(Ay)
End Sub

Private Sub ZZ_CntDryzAy()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = CntDryzAy(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
Stop
'AssEqDry Act, Exp
End Sub
