Attribute VB_Name = "MVb_Ay_Map_Align"
Option Explicit
Function FmtAyNTerm(Ay, N%) As String()
Dim W%(), L
W = WdtAyNTermAy(N, Ay)
For Each L In Itr(Ay)
    PushI FmtAyNTerm, FmtAyNTerm1(L, W)
Next
End Function

Private Function FmtAyNTerm1$(A, W%())
Dim Ay$(), J%, N%, O$(), I
N = Sz(W)
Ay = SyzNTermRst(A, N)
If Sz(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, AlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
FmtAyNTerm1 = RTrim(JnSpc(O))
End Function

Private Function WdtAyNTermAy(NTerm%, Ay) As Integer()
If Sz(Ay) = 0 Then Exit Function
Dim O%(), W%(), L
ReDim O(NTerm - 1)
For Each L In Ay
    W = WdtAyNTermLin(NTerm, L)
    O = WdtAyAB(O, W)
Next
WdtAyNTermAy = O
End Function

Private Function WdtAyNTermLin(N%, Lin) As Integer()
Dim T
For Each T In FstNTerm(Lin, N)
    PushI WdtAyNTermLin, Len(T)
Next
End Function
Private Function WdtAyAB(A%(), B%()) As Integer()
Dim O%(), J%, I
O = A
For Each I In B
    If I > O(J) Then O(J) = I
    J = J + 1
Next
WdtAyAB = O
End Function

Function FmtAyAtChr(Ay, AtChr$) As String()
Dim T1$(), Rst$(), I, P%
For Each I In Itr(Ay)
    P = InStr(I, AtChr)
    If P = 0 Then
        PushI T1, ""
        PushI Rst, I
    Else
        PushI T1, Left(I, P)
        PushI Rst, Mid(I, P + 1)
    End If
Next
Dim J&
T1 = FmtAyR(T1)
For Each I In Itr(T1)
    PushI FmtAyAtChr, I & Rst(J)
    J = J + 1
Next
End Function

Function FmtAyAtDot(A) As String()
FmtAyAtDot = FmtAyAtChr(A, ".")
End Function
Function FmtAyColon(ColonAy) As String()
FmtAyColon = FmtDry(DryColonAy(ColonAy))
End Function

Function FmtAyDot(DotAy) As String()
FmtAyDot = FmtDry(DryDotAy(DotAy), Fmt:=eSpcSep)
End Function

Function FmtAyT1(A) As String()
FmtAyT1 = FmtAyNTerm(A, 1)
End Function

Function FmtAy2T(A) As String()
FmtAy2T = FmtAyNTerm(A, 2)
End Function

Function FmtAy3T(A$()) As String()
FmtAy3T = FmtAyNTerm(A, 3)
End Function

Function FmtAy4T(A$()) As String()
FmtAy4T = FmtAyNTerm(A, 4)
End Function

Function FmtAyAlign(Ay) As String()
Dim W%: W = WdtzAy(Ay) + 1
Dim I
For Each I In Itr(Ay)
    Push FmtAyAlign, AlignL(I, W)
Next
End Function

Function FmtAyR(Ay) As String()
Dim W%: W = WdtzAy(Ay)
Dim I
For Each I In Itr(Ay)
    Push FmtAyR, AlignR(I, W)
Next
End Function

Private Sub Z_FmtAy2T()
Dim Ly$()
Ly = Sy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtAy2T(Ly)
    C
    Return
End Sub
Private Sub Z_FmtAy3T()
Dim Ly$()
Ly = Sy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C   D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtAy3T(Ly)
    C
    Return
End Sub




Private Sub Z()
Z_FmtAy2T
Z_FmtAy3T
MVb_Align_Ay:
End Sub
