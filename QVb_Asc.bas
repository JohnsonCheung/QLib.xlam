Attribute VB_Name = "QVb_Asc"
Option Explicit
Private Const CMod$ = "MVb_Asc."
Private Const Asm$ = "QVb"

Function AscAt%(S, Pos)
AscAt = Asc(Mid(S, Pos, 1))
End Function

Function IsAscCrLf(Asc%)
IsAscCrLf = (Asc = 13) Or (Asc = 10)
End Function

