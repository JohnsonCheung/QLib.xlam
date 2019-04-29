Attribute VB_Name = "MVb_Asc"
Option Explicit

Function AscAt%(S, Pos)
AscAt = Asc(Mid(S, Pos, 1))
End Function

Function IsAscCrLf(Asc%)
IsAscCrLf = (Asc = 13) Or (Asc = 10)
End Function

