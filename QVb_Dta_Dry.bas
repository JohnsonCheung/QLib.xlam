Attribute VB_Name = "QVb_Dta_Dry"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dta."
Private Const Asm$ = "QVb"

Function IsEqDry(A(), B()) As Boolean
IsEqDry = IsEqAy(A, B)
End Function

Private Sub ZZ()
Dim A()
IsEqDry A, A
End Sub

