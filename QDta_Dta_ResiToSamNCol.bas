Attribute VB_Name = "QDta_Dta_ResiToSamNCol"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Dy_ReSzToSamColCnt."
Private Const Asm$ = "QDta"
Function ResiToSamNCol(A As Drs) As Drs
Dim N%: N = NColzDrs(A)
ResiToSamNCol = A
ResiToSamNCol.Dy = ResiToNCol(A.Dy, N - 1)
End Function
Function ResiToSamNColzDy(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
'ResiToSamNColzDy = ResiToUCol(O, NColzDy(Dy) - 1)
End Function

Private Function ResiToNCol(Dy(), U%) As Variant()
Dim Dr, O(), J&
O = Dy
For Each Dr In Itr(O)
    If UB(Dr) <> U Then
        ReDim Preserve Dr(U)
        O(J) = Dr
    End If
    J = J + 1
Next
ResiToNCol = O
End Function
