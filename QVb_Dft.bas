Attribute VB_Name = "QVb_Dft"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dft."
Private Const Asm$ = "QVb"
Function Dft(V, DftV)
If IsEmp(V) Then
   Dft = DftV
Else
   Dft = V
End If
End Function

Function DftStr$(Str, Dft)
DftStr = IIf(Str = "", Dft, Str)
End Function

