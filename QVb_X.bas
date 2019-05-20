Attribute VB_Name = "QVb_X"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_X."
Private Const Asm$ = "QVb"
Public XX$()
Sub X(S)
PushI XX, S
End Sub
Sub XAy(Ay)
PushIAy XX, Ay
End Sub

Function Y(A)
Stop
End Function
