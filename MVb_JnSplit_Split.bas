Attribute VB_Name = "MVb_JnSplit_Split"
Option Explicit
Function SplitComma(A) As String()
SplitComma = Split(A, ",")
End Function

Function SplitCommaSpc(A) As String()
SplitCommaSpc = Split(A, ", ")
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(Replace(A, vbCr, ""), vbLf)
End Function

Function SplitTab(A) As String()
SplitTab = Split(A, vbTab)
End Function

Function SplitDot(A) As String()
SplitDot = Split(A, ".")
End Function

Function SplitColon(A) As String()
SplitColon = Split(A, ":")
End Function

Function SplitSemi(A) As String()
SplitSemi = Split(A, ";")
End Function

Function SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Function

Function SplitSsl(A) As String()
SplitSsl = Split(RplDblSpc(Trim(A)), " ")
End Function

Function SplitVbar(A) As String()
SplitVbar = AyTrim(Split(A, "|"))
End Function
