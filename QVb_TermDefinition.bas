Attribute VB_Name = "QVb_TermDefinition"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_TermDefinition."
Private Const Asm$ = "QVb"

Function DefzCml() As String()
Erase XX
X "Cml Cml is a string contains only letter-and-digit, (no underscore)"
X "Cml CmlSy is breaking Cml in Ay of (CmlFstTerm + N-CmlTerm) "
X "Cml CmlTerm is One-UCase + N-(LCase-or-Digit)"
X "Cml CmlFstTerm is CmlTerm or (Lcase + N-(LCase-or-Digit))"
DefzCml = FmtSy3Term(XX)
Erase XX
End Function
