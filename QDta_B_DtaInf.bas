Attribute VB_Name = "QDta_B_DtaInf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Col_Get."
Private Const Asm$ = "QDta"
Public Const vbFldSep$ = ""
Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function ValzColEq(A As Drs, SelC$, Col$, Eq)
Dim Dr, I&
I = IxzAy(A.Fny, Col)
For Each Dr In Itr(A.Dry)
    If Dr(I) = Eq Then ValzColEq = Dr(IxzAy(A.Fny, SelC))
Next
End Function

Function WdtzCol%(A As Drs, C$)
WdtzCol = WdtzAy(StrColzDrs(A, C))
End Function
Function StrCol(A As Drs, C) As String()
StrCol = StrColzDry(A.Dry, IxzAy(A.Fny, C))
End Function

Function StrColzDrs(A As Drs, C) As String()
StrColzDrs = StrColzDry(A.Dry, IxzAy(A.Fny, C))
End Function
Function BoolColzDrs(A As Drs, C) As Boolean()
BoolColzDrs = BoolColzDry(A.Dry, IxzAy(A.Fny, C))
End Function

Function FstCol(A As Drs) As Variant()
FstCol = AvzDryC(A.Dry, 0)
End Function

Function StrColzDrsFstCol(A As Drs) As String()
StrColzDrsFstCol = StrColzDry(A.Dry, 0)
End Function

Function StrColzColEqSel(A As Drs, Col$, V, ColNm$) As String()
Dim B As Drs
B = ColEqSel(A, Col, V, ColNm)
StrColzColEqSel = StrColzDrs(B, ColNm)
End Function
Function LyzDryCC(Dry(), CCIxy&(), Optional FldSep$ = vbFldSep) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI LyzDryCC, Jn(AyeIxy(Dr, CCIxy), FldSep)
Next
End Function

Function SqzAyV(Ay) As Variant()
Dim J&, O()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = Ay(J - 1)
Next
SqzAyV = O
End Function

Function SqzAyH(Ay) As Variant()
Dim J&, O()
Dim N&: N = Si(Ay)
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = Ay(J - 1)
Next
SqzAyH = O
End Function

Function SqzDry(Dry()) As Variant()
SqzDry = SqzDrySkip(Dry, 0)
End Function

Function StrColzDry(Dry(), C) As String()
StrColzDry = IntozDryC(EmpSy, Dry, C)
End Function
Function BoolColzDry(Dry(), C&) As Boolean()
BoolColzDry = IntozDryC(EmpBoolAy, Dry, C)
End Function

Function SqzDrySkip(Dry(), Optional SkipNRow& = 1)
SqzDrySkip = SqzDry(CvAv(AySkip(Dry, SkipNRow)))
End Function

Function IntAyzDrsC(A As Drs, C) As Integer()
IntAyzDrsC = IntAyzDryC(A.Dry, IxzAy(A.Fny, C))
End Function

Function IntAyzDryC(Dry(), C) As Integer()
IntAyzDryC = IntozDryC(EmpIntAy, Dry, C)
End Function

Function ColzDry(Dry(), C) As Variant()
ColzDry = IntozDryC(EmpAv(), Dry, C)
End Function

Function AvzDryC(Dry(), C) As Variant()
AvzDryC = IntozDryC(EmpAv, Dry, C)
End Function
Function IntozDryC(Into, Dry(), C)
Dim O, J&, Dr, U&
O = Into
U = UB(Dry)
O = Resi(O, U)
For Each Dr In Itr(Dry)
    If UB(Dr) >= C Then
        O(J) = Dr(C)
    End If
    J = J + 1
Next
IntozDryC = O
End Function

Function HasDryCCEqVV(Dry(), C1, C2, V1, V2) As Boolean
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C1) = V1 Then
        If Dr(C2) = V2 Then
            HasDryCCEqVV = True
            Exit Function
        End If
    End If
Next
End Function
Function IsSamNCol(A As Drs, NCol%) As Boolean
Dim Dr
For Each Dr In Itr(A.Dry)
    If Si(Dr) = NCol Then Exit Function
Next
IsSamNCol = True
End Function
Function ResiDrs(A As Drs, NCol%) As Drs
If IsSamNCol(A, NCol) Then ResiDrs = A: Exit Function
Dim O As Drs, U%, Dr, J%
U = NCol - 1
For J = 0 To UB(O.Dry)
    Dr = O.Dry(J)
    ReDim Preserve Dr(U)
    O.Dry(J) = Dr
Next
End Function

