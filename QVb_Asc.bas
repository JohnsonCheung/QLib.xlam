Attribute VB_Name = "QVb_Asc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Asc."
Private Const Asm$ = "QVb"

Function AscAt%(S, At)
AscAt = Asc(Mid(S, At, 1))
End Function

Function IsStrAtSpcCrLf(S, At) As Boolean
IsStrAtSpcCrLf = IsAscSpcCrLf(AscAt(S, At))
End Function

Function IsAscSpcCrLf(Asc%)
Select Case True
Case Asc = 13, Asc = 10, Asc = 32: IsAscSpcCrLf = True
End Select
End Function

