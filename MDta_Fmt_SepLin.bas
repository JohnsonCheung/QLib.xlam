Attribute VB_Name = "MDta_Fmt_SepLin"
Option Explicit
Function SepLin$(W%(), Sep$)
SepLin = SepLinzSepDr(SepDr(W), Sep)
End Function
Function SepDr(W%()) As String()
Dim I
For Each I In W
    Push SepDr, Dup("-", I)
Next
End Function

Function SepLinzSepDr$(SepDr$(), Sep$)
SepLinzSepDr = "|" & Join(SepDr, Sep) & "|"
End Function

Function SepChrzDryFmt$(A As eDryFmt)
Dim O$
Select Case A
Case eDryFmt.eSpcSep: O = " "
Case eDryFmt.eSpcSep: O = "|"
Case Else: Thw CSub, "Invalid eDryFmt", "eDryFmt", A
End Select
SepChrzDryFmt = O
End Function
