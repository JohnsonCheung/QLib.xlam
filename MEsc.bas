Attribute VB_Name = "MEsc"
Function HexzChr$(C$)
If Len(C) <> 1 Then Thw CSub, "C should have len=1", "C Len", C, Len(C)
Dim A%: A = Asc(C)
If A < 16 Then
    HexzChr = "0" & Hex(A)
Else
    HexzChr = Hex(A)
End If
End Function
Function PerHex$(C$)
PerHex = "%" & HexzChr(C)
End Function
Function EscChr$(S$, C$)
EscChr = EscAsc(S, Asc(C))
End Function
Function EscAsc$(S$, A%)
EscChr = Replace(S, "%" & Hex2(A))
End Function
Function EscSqBkt$(S$)
EscSqBkt = EscChrLis(S, "[]")
End Function
Function EscChrLis$(S$, ChrLis$)
Dim O$: O = S
For J = 1 To Len(ChrLis)
    O = EscChr(O, Mid(ChrLis, J, 1))
Next
EscChrLis = O
End Function

