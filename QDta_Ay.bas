Attribute VB_Name = "QDta_Ay"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Ay."
Private Const Asm$ = "QDta"

Function DtzAy(Ay, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim Dry(), J&
For J = 0 To UB(Ay)
    PushI Dry, Array(Ay(J))
Next
DtzAy = Dt(DtNm, Sy(FldNm), Dry)
End Function

Function DryGpCntzAy(Ay) As Variant()
If Si(Ay) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AywDup(Ay)
For Each X In Itr(Dup)
    Cnt = AyItmCnt(Ay, X)
    Push O, Array(X, AyItmCnt(Ay, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
DryGpCntzAy = O
End Function

Function DryGpCntzAyWhDup(A) As Variant()
DryGpCntzAyWhDup = DrywColGt(DryGpCntzAy(A), 1, 1)
End Function
Sub BrwDryGpCntzAy(Ay)
Brw LyzDryGpCntzAy(Ay)
End Sub

Function LyzDryGpCntzAy(Ay) As String()
LyzDryGpCntzAy = LyzDry(DryGpCntzAy(Ay), Fmt:=EiSSFmt)
End Function

Private Sub Z_LyzDryGpCntzAy()
Dim Ay()
Brw LyzDryGpCntzAy(Ay)
End Sub

Private Sub Z_CntDryzAy()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = CntDryzAy(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
Stop
'AssEqDry Act, Exp
End Sub
