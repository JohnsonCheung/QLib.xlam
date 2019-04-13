Attribute VB_Name = "MVb_Str_Rpl"
Option Explicit
Private Sub ZZ_RplBet()
Dim A$, Exp$, By$, s1$, s2$
s1 = "Data Source="
s2 = ";"
A = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(A, By, s1, s2)
Debug.Assert Exp = Act
Return
End Sub

Private Sub ZZ_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub
Function RmvCr$(A)
RmvCr = Replace(A, vbCr, "")
End Function

Function RplCr$(A)
RplCr = Replace(A, vbCr, " ")
End Function
Function RplLf$(A)
RplLf = Replace(A, vbLf, " ")
End Function
Function RplVbl$(Vbl)
RplVbl = RplVBar(Vbl)
End Function
Function RplVBar$(Vbl)
RplVBar = Replace(Vbl, "|", vbCrLf)
End Function
Function RplBet$(A, By$, s1$, s2$)
Dim P1%, P2%, B$, C$
P1 = InStr(A, s1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(s1), CStr(A), s2)
If P2 = 0 Then Stop
B = Left(A, P1 + Len(s1) - 1)
C = Mid(A, P2 + Len(s2) - 1)
RplBet = B & By & C
End Function

Function RplDblSpc$(A)
Dim O$: O = Trim(A)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplFstChr$(A, By$)
RplFstChr = By & RmvFstChr(A)
End Function

Function RplPfx(A, FmPfx, ToPfx)
RplPfx = ToPfx & RmvPfx(A, FmPfx)
End Function

Private Sub Z_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub

Function RplPun$(A)
Dim O$(), J&, L&, C$
L = Len(A)
If L = 0 Then Exit Function
ReDim O(L - 1)
For J = 1 To L
    C = Mid(A, J, 1)
    If IsPun(C) Then
        O(J - 1) = " "
    Else
        O(J - 1) = C
    End If
Next
RplPun = Join(O, "")
End Function

Function RplQ$(A, By)
RplQ = Replace(A, "?", By)
End Function

Private Sub Z_RplBet()
Dim A$, Exp$, By$, s1$, s2$
s1 = "Data Source="
s2 = ";"
A = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(A, By, s1, s2)
Debug.Assert Exp = Act
Return
End Sub


Private Sub Z()
Z_RplBet
Z_RplPfx
MVb_Str_Rpl:
End Sub
