Attribute VB_Name = "MDta_Fmt_SepLin"
Option Explicit
Function SepChr$(A As eDryFmt)
Select Case A
Case eDryFmt.eSpcSep: SepChr = "-"
Case eDryFmt.eVbarSep: SepChr = "|"
Case Else: ThwPmEr "DryFmt", CSub
End Select
End Function
Function SepLin$(W%(), Fmt As eDryFmt)
SepLin = SepLinzSepDr(SepDr(W), Fmt)
End Function
Function SepDr(W%()) As String()
Dim I
For Each I In W
    Push SepDr, Dup("-", I)
Next
End Function

Function SepLinzSepDr$(SepDr$(), Fmt As eDryFmt)
SepLinzSepDr = "|-" & Join(SepDr, SepChr(Fmt)) & "-|"
End Function
