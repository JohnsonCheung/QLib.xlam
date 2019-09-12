Attribute VB_Name = "MxDtaS12"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaS12."

Function DrszS12s(A As S12s) As Drs
DrszS12s = DrszFF("S1 S2", DyzS12s(A))
End Function

Private Function AvzS12(A As S12) As Variant()
AvzS12 = Array(A.S1, A.S2)
End Function

Function DyzAyab(A, B) As Variant()
ThwIf_DifSi A, B, CSub
Dim I, J&: For Each I In Itr(A)
    PushI DyzAyab, Array(I, B(J))
    J = J + 1
Next
End Function

Function DyzS12s(A As S12s) As Variant()
Dim J&
For J = 0 To A.N - 1
    PushI DyzS12s, AvzS12(A.Ay(J))
Next
End Function
