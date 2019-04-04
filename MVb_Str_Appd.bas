Attribute VB_Name = "MVb_Str_Appd"
Option Explicit

Function ApdCrLf$(A)
If A = "" Then Exit Function
ApdCrLf = A & vbCrLf
End Function
Function PpdSpc$(A)
If A = "" Then Exit Function
PpdSpc = " " & A
End Function
Function Apd$(A, Sfx, Optional Sep$ = "")
If A = "" Then Apd = A: Exit Function
Apd = A & Sep & Sfx
End Function

Function Ppd$(A, Pfx, Optional Sep$ = "")
If A = "" Then Ppd = A: Exit Function
Ppd = Pfx & Sep & A
End Function

