Attribute VB_Name = "QVb_Str_Rpl"
Option Explicit
Private Const CMod$ = "MVb_Str_Rpl."
Private Const Asm$ = "QVb"
Private Sub ZZ_RplBet()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub

Private Sub ZZ_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub
Function RmvCr$(S$)
RmvCr = Replace(S, vbCr, "")
End Function

Function RplCr$(S$)
RplCr = Replace(S, vbCr, " ")
End Function
Function RplLf$(S$)
RplLf = Replace(S, vbLf, " ")
End Function
Function RplVbl$(S$)
RplVbl = RplVBar(S)
End Function
Function RplVBar$(S$)
RplVBar = Replace(S, "|", vbCrLf)
End Function
Function RplBet$(S$, By$, S1$, S2$)
Dim P1%, P2%, B$, C$
P1 = InStr(S, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), CStr(S), S2)
If P2 = 0 Then Stop
B = Left(S, P1 + Len(S1) - 1)
C = Mid(S, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function

Function RplDblSpc$(S$)
Dim O$: O = Trim(S)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplFstChr$(S$, By$)
RplFstChr = By & RmvFstChr(S)
End Function

Function RplPfx(S$, Fm$, ToPfx$)
If HasPfx(S, Fm) Then
    RplPfx = ToPfx & RmvPfx(S, Fm)
Else
    RplPfx = S
End If
End Function

Private Sub Z_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub

Function RplPun$(S$)
Dim O$(), J&, L&, C$
L = Len(S)
If L = 0 Then Exit Function
ReDim O(L - 1)
For J = 1 To L
    C = Mid(S, J, 1)
    If IsPun(C) Then
        O(J - 1) = " "
    Else
        O(J - 1) = C
    End If
Next
RplPun = Join(O, "")
End Function

Function RplQ$(S, By)
RplQ = Replace(S, "?", By)
End Function

Private Sub Z_RplBet()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub


Private Sub Z()
Z_RplBet
Z_RplPfx
MVb_Str_Rpl:
End Sub
