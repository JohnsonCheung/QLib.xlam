Attribute VB_Name = "QVb_Dta_Dry"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dta."
Private Const Asm$ = "QVb"

Function IsEqDy(A(), B()) As Boolean
IsEqDy = IsEqAy(A, B)
End Function

Private Sub Z()
Dim A()
IsEqDy A, A
End Sub

