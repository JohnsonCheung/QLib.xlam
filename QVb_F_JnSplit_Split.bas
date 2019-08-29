Attribute VB_Name = "QVb_F_JnSplit_Split"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_JnSplit_Split."
Private Const Asm$ = "QVb"
Function SplitComma(S) As String()
SplitComma = Split(S, ",")
End Function

Function SplitCommaSpc(S) As String()
SplitCommaSpc = Split(S, ", ")
End Function
Function Ly(Lines) As String()
Ly = SplitCrLf(Lines)
End Function

Function SplitCrLf(S) As String()
SplitCrLf = Split(Replace(S, vbCr, ""), vbLf)
End Function

Function SplitTab(S) As String()
SplitTab = Split(S, vbTab)
End Function

Function SplitDot(S) As String()
SplitDot = Split(S, ".")
End Function

Function SplitColon(S) As String()
SplitColon = Split(S, ":")
End Function

Function SplitSemi(S) As String()
SplitSemi = Split(S, ";")
End Function

Function SplitSpc(S) As String()
SplitSpc = Split(S, " ")
End Function

Function SplitSsl(S) As String()
SplitSsl = Split(RplDblSpc(Trim(S)), " ")
End Function

Function SplitVBar(S) As String()
SplitVBar = AyTrim(CvSy(Split(S, "|")))
End Function

Function LyzLinesAy(LinesAy$()) As String()
Dim L: For Each L In Itr(LinesAy)
    PushIAy LyzLinesAy, SplitCrLf(L)
Next
End Function


'
