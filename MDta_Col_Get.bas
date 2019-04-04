Attribute VB_Name = "MDta_Col_Get"
Option Explicit
Public Const vbFldSep$ = ""
Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function StrColzDrs(A As Drs, ColNm$) As String()
StrColzDrs = StrColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function DrLinAy(Dry(), Optional CC, Optional FldSep$ = vbFldSep) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrLinAy, DrLin(Dr, CC, FldSep)
Next
End Function

Function DrLin$(Dr, Optional CC, Optional FldSep$ = vbFldSep)
If IsMissing(CC) Then
    DrLin = Jn(Dr, FldSep)
Else
    DrLin = Jn(AywIxAy(Dr, IxAyzCC(CC)), FldSep)
End If
End Function

Function SqzDry(A()) As Variant()
SqzDry = SqzDrySkip(A, 0)
End Function

Function StrColzDry(Dry(), C) As String()
StrColzDry = IntoColzDry(EmpSy, Dry, C)
End Function

Function SqzDrySkip(A(), SkipNRow%)
Dim O(), C%, R&, Dr
Dim NC%, NR&
NC = NColzDry(A)
NR = Si(A) + SkipNRow
ReDim O(1 To NR, 1 To NC)
Dim DryIx&
For R = SkipNRow + 1 To NR
    Dr = A(DryIx)
    SetSqrzDr O, R, Dr
    DryIx = DryIx + 1
Next
SqzDrySkip = O
End Function

Function IntAyDryC(A(), C) As Integer()
IntAyDryC = IntoColzDry(EmpIntAy, A, C)
End Function

Function ColzDry(Dry(), C) As Variant()
ColzDry = IntoColzDry(EmpAv(), Dry, C)
End Function

Function IntoColzDry(Into, Dry(), C)
Dim O, J&, Dr, U&
O = Into
U = UB(Dry)
ReSumSiU O, U
For Each Dr In Itr(Dry)
    If UB(Dr) >= C Then
        O(J) = Dr(C)
    End If
    J = J + 1
Next
IntoColzDry = O
End Function


