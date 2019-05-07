Attribute VB_Name = "QDta_Col_DtaInf"
Option Explicit
Private Const CMod$ = "MDta_Col_Get."
Private Const Asm$ = "QDta"
Public Const vbFldSep$ = ""
Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function StrColzDrs(A As Drs, ColNm$) As String()
StrColzDrs = StrColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function LyzDryCC(Dry(), CCIxAy&(), Optional FldSep$ = vbFldSep) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI LyzDryCC, Jn(AyeIxAy(Dr, CCIxAy), FldSep)
Next
End Function

Function SqzDry(Dry()) As Variant()
SqzDry = SqzDrySkip(Dry, 0)
End Function

Function StrColzDry(Dry(), C&) As String()
StrColzDry = IntozDryc(EmpSy, Dry, C)
End Function
Function SqzDrySkip(Dry(), Optional SkipNRow& = 1)
SqzDrySkip = SqzDry(CvAv(AySkip(Dry, SkipNRow)))
End Function

Function IntAyzDryc(Dry(), C&) As Integer()
IntAyzDryc = IntozDryc(EmpIntAy, Dry, C)
End Function

Function ColzDry(Dry(), C&) As Variant()
ColzDry = IntozDryc(EmpAv(), Dry, C)
End Function

Function IntozDryc(Into, Dry(), C&)
Dim O, J&, Drv, U&
O = Into
U = UB(Dry)
O = Resi(O, U)
For Each Drv In Itr(Dry)
    If UB(Drv) >= C Then
        O(J) = Drv(C)
    End If
    J = J + 1
Next
IntozDryc = O
End Function
