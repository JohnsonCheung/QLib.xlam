Attribute VB_Name = "QVb_Dic_Diw"
Option Explicit
Option Compare Text
Function DiwAy(D As Dictionary, Ay) As Dictionary
Set DiwAy = New Dictionary
Dim K: For Each K In D.Keys
    If HasEle(Ay, K) Then DiwAy.Add K, D(K)
Next
End Function


'
