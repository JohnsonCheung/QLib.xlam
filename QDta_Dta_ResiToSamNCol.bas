Attribute VB_Name = "QDta_Dta_ResiToSamNCol"
Option Explicit
Private Const CMod$ = "MDta_Dry_ReSzToSamColCnt."
Private Const Asm$ = "QDta"
Function ResiToSamNCol(A As Drs) As Drs
Dim N%: N = NColzDrs(A)
ResiToSamNCol = A
ResiToSamNCol.Dry = ResiToNCol(A.Dry, N - 1)
End Function
Function ResiToSamNColzDry(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
'ResiToSamNColzDry = ResiToUCol(O, NColzDry(Dry) - 1)
End Function

Private Function ResiToNCol(Dry(), U%) As Variant()
Dim Dr, O(), J&
O = Dry
For Each Dr In Itr(O)
    If UB(Dr) <> U Then
        ReDim Preserve Dr(U)
        O(J) = Dr
    End If
    J = J + 1
Next
ResiToNCol = O
End Function
