Attribute VB_Name = "MVb_TermDefinition"
Option Explicit

Function DefzCml() As String()
Erase xx
X "Cml Cml is a string contains only letter-and-digit, (no underscore)"
X "Cml CmlAy is breaking Cml in Ay of (CmlFstTerm + N-CmlTerm) "
X "Cml CmlTerm is One-UCase + N-(LCase-or-Digit)"
X "Cml CmlFstTerm is CmlTerm or (Lcase + N-(LCase-or-Digit))"
DefzCml = FmtAyT3(xx)
Erase xx
End Function
