Attribute VB_Name = "MVb_Str_Appd"
Option Explicit

Function AppdCrLf$(A)
If A = "" Then Exit Function
AppdCrLf = A & vbCrLf
End Function
Function PrepSpc$(A)
If A = "" Then Exit Function
PrepSpc = " " & A
End Function
Function Appd$(A, Sfx, Optional Sep$ = "")
If A = "" Then Appd = A: Exit Function
Appd = A & Sep & Sfx
End Function

Function Prep$(A, Pfx, Optional Sep$ = "")
If A = "" Then Prep = A: Exit Function
Prep = Pfx & Sep & A
End Function

