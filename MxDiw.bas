Attribute VB_Name = "MxDiw"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxDiw."
Function DiwAy(D As Dictionary, Ay) As Dictionary
Set DiwAy = New Dictionary
Dim K: For Each K In D.Keys
    If HasEle(Ay, K) Then DiwAy.Add K, D(K)
Next
End Function