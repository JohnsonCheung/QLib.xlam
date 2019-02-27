Attribute VB_Name = "MVb_X"
Option Explicit
Public XX$()
Sub X(S$)
Push XX, S
End Sub
Sub X0(S$)
If Sz(XX) = 0 Then PushI XX, S: Exit Sub
X1 S
End Sub
Sub X1(S$)
Dim U&: U = UB(XX)
XX(U) = XX(U) & S
End Sub
