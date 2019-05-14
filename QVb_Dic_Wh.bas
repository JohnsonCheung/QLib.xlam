Attribute VB_Name = "QVb_Dic_Wh"
Option Explicit
Private Const CMod$ = "MVb_Dic_Wh."
Private Const Asm$ = "QVb"
Function DicwKK(A As Dictionary, KK) As Dictionary
Set DicwKK = New Dictionary
Dim K
For Each K In TermAy(KK)
    If A.Exists(K) Then
        DicwKK.Add K, A(K)
    End If
Next
End Function
