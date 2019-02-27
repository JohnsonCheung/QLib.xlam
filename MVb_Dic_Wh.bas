Attribute VB_Name = "MVb_Dic_Wh"
Option Explicit
Function DicwKK(A As Dictionary, KK) As Dictionary
Set DicwKK = New Dictionary
Dim K
For Each K In TermAy(KK)
    If A.Exists(K) Then
        DicwKK.Add K, A(K)
    End If
Next
End Function
