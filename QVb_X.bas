Attribute VB_Name = "QVb_X"
Option Explicit
Private Const CMod$ = "MVb_X."
Private Const Asm$ = "QVb"
Public XX$()
Sub X(S$)
Push XX, S
End Sub
Sub X0(S$)
If Si(XX) = 0 Then PushI XX, S: Exit Sub
X1 S
End Sub
Sub X1(S$)
Dim U&: U = UB(XX)
XX(U) = XX(U) & S
End Sub
