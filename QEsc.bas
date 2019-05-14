Attribute VB_Name = "QEsc"
Option Explicit
Private Const CMod$ = "MEsc."
Private Const Asm$ = "Q"
Function Hex2zAsc$(Asc%)
If Asc < 16 Then
    Hex2zAsc = "0" & Hex(Asc)
Else
    Hex2zAsc = Hex(Asc)
End If
End Function
Function Hex2$(C$)
If Len(C) <> 1 Then Thw CSub, "C should have len=1", "C Len", C, Len(C)
Hex2 = Hex2zAsc(Asc(C))
End Function
Function PerHex2$(C$)
PerHex2 = "%" & Hex2(C)
End Function
Function EscChr$(S, C$)
EscChr = EscAsc(S, Asc(C))
End Function
Function EscAsc$(S, A%) 'Escaping the AscChr-A% in S$ as %HH
EscAsc = Replace(S, Chr(A), "%" & Hex2zAsc(A))
End Function
Function EscSqBkt$(S)
EscSqBkt = EscChrLis(S, "[]")
End Function
Function EscChrLis$(S, ChrLis$)
Dim O$, J: O = S
For J = 1 To Len(ChrLis)
    O = EscChr(O, Mid(ChrLis, J, 1))
Next
EscChrLis = O
End Function

