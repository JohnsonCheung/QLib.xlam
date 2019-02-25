Attribute VB_Name = "MDta_Col_Get"
Option Explicit
Function ColzDrs(A As Drs, ColNm$) As Variant()
ColzDrs = ColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function
Function StrColzDrs(A As Drs, ColNm$) As String()
StrColzDrs = StrColzDry(A.Dry, IxzAy(A.Fny, ColNm))
End Function

Function SyzDry(Dry(), C) As String()
SyzDry = StrColzDry(Dry, C)
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
NC = NColDry(A)
NR = Sz(A) + SkipNRow
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
ReszAyU O, U
For Each Dr In Itr(Dry)
    If UB(Dr) >= C Then
        O(J) = Dr(C)
    End If
    J = J + 1
Next
IntoColzDry = O
End Function


