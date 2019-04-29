Attribute VB_Name = "MVb_Dft"
Option Explicit
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

