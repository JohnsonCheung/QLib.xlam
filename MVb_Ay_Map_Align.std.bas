Attribute VB_Name = "MVb_Ay_Map_Align"
Option Explicit
Function AyAlignNTerm(Ay, N%) As String()
Dim W%(), L
W = WdtAyNTermAy(N, Ay)
For Each L In Itr(Ay)
    PushI AyAlignNTerm, AyAlignNTerm1(L, W)
Next
End Function

Function AyAlignT1(A) As String()
Dim T1$(), Rest$()
    AyAsgT1AyRestAy A, T1, Rest
T1 = AyAlignL(T1)
AyAlignT1 = JnAyabSpc(T1, Rest)
End Function

Private Function AyAlignNTerm1$(A, W%())
Dim Ay$(), J%, N%, O$(), I
N = Sz(W)
Ay = SyNTermRst(A, N)
If Sz(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, AlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
AyAlignNTerm1 = RTrim(JnSpc(O))
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

Function AyAlignAtChr(Ay, AtChr$) As String()
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
T1 = AyAlignR(T1)
For Each I In Itr(T1)
    PushI AyAlignAtChr, I & Rst(J)
    J = J + 1
Next
End Function

Function AyAlignAtDot(A) As String()
AyAlignAtDot = AyAlignAtChr(A, ".")
End Function
Function AyAlignColon(ColonAy) As String()
AyAlignColon = FmtDry(DryColonAy(ColonAy))
End Function

Function AyAlignDot(DotAy) As String()
AyAlignDot = FmtDry(DryDotAy(DotAy), Fmt:=eSpcSep)
End Function
Function FmtAyT1(A) As String()
FmtAyT1 = AyAlignNTerm(A, 1)
End Function

Function AyAlign2T(A) As String()
AyAlign2T = AyAlignNTerm(A, 2)
End Function

Function AyAlign3T(A$()) As String()
AyAlign3T = AyAlignNTerm(A, 3)
End Function

Function AyAlign4T(A$()) As String()
AyAlign4T = AyAlignNTerm(A, 4)
End Function

Function AyAlignL(Ay) As String()
Dim W%: W = WdtzAy(Ay) + 1
Dim I
For Each I In Itr(Ay)
    Push AyAlignL, AlignL(I, W)
Next
End Function

Function AyAlignR(Ay) As String()
Dim W%: W = WdtzAy(Ay)
Dim I
For Each I In Itr(Ay)
    Push AyAlignR, AlignR(I, W)
Next
End Function

Private Sub Z_AyAlign2T()
Dim Ly$()
Ly = Sy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AyAlign2T(Ly)
    C
    Return
End Sub
Private Sub Z_AyAlign3T()
Dim Ly$()
Ly = Sy("AAA B C D", "A BBB CCC")
Ept = Sy("AAA B   C   D", _
         "A   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AyAlign3T(Ly)
    C
    Return
End Sub




Private Sub Z()
Z_AyAlign2T
Z_AyAlign3T
MVb_Align_Ay:
End Sub
