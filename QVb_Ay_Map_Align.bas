Attribute VB_Name = "QVb_Ay_Map_Align"
Option Explicit
Private Const CMod$ = "MVb_Ay_Map_Align."
Private Const Asm$ = "QVb"
Enum EmAlign
    EiLeft
    EiRight
End Enum
Function FmtAyNTerm(Ay, N%) As String()
Dim W%(), L
W = WdtAyNTermSy(N, Ay)
For Each L In Itr(Ay)
    PushI FmtAyNTerm, FmtAyNTerm1(L, W)
Next
End Function

Private Function FmtAyNTerm1$(Sy, W%())
Dim Ay$(), J%, N%, O$(), I
N = Si(W)
Ay = SyzNTermRst(Sy, N)
If Si(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, AlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
FmtAyNTerm1 = RTrim(JnSpc(O))
End Function

Private Function WdtAyNTermSy(NTerm%, Ay) As Integer()
If Si(Ay) = 0 Then Exit Function
Dim O%(), W%(), L
ReDim O(NTerm - 1)
For Each L In Ay
    W = WdtAyNTermLin(NTerm, L)
    O = WdtAyab(O, W)
Next
WdtAyNTermSy = O
End Function

Private Function WdtAyNTermLin(N%, Lin) As Integer()
Dim T
For Each T In FstNTerm(Lin, N)
    PushI WdtAyNTermLin, Len(T)
Next
End Function
Private Function WdtAyab(Sy%(), B%()) As Integer()
Dim O%(), J%, I
O = Sy
For Each I In B
    If I > O(J) Then O(J) = I
    J = J + 1
Next
WdtAyab = O
End Function
Function S1S2zAtChr(S$, AtChr$, Optional IfNoAtChr As EmAlign) As S1S2
Dim P%: P = InStr(S, AtChr)
Select Case True
Case P = 0 And IfNoAtChr = EiLeft:  S1S2zAtChr = S1S2(S, "")
Case P = 0 And IfNoAtChr = EiRight: S1S2zAtChr = S1S2("", S)
Case Else:                          S1S2zAtChr = S1S2(Left(S, P - 1), Mid(S, P))
End Select
End Function

Function S1S2szSyAtChr(Sy$(), AtChr$, Optional IfNotAtChr As EmAlign) As S1S2s
Dim I
For Each I In Itr(Sy)
    PushS1S2 S1S2szSyAtChr, S1S2zAtChr(CStr(I), AtChr, EiLeft)
Next
End Function
Function AlignAtChr(Sy$(), AtChr$, Optional IfNoAtChr As EmAlign) As String()
AlignAtChr = FmtS1S2s(S1S2szSyAtChr(Sy, AtChr))
End Function

Function AlignAtDot(Sy$(), Optional IfNoDt As EmAlign) As String()
AlignAtDot = AlignAtChr(Sy, ".")
End Function

Sub BrwDotLy(DotLy$())
Brw FmtAyDot(DotLy)
End Sub

Function FmtAyDot(DotLy$()) As String()
FmtAyDot = FmtDryAsSpcSep(DryzDotLy(DotLy))
End Function
Function FmtAyDot1(DotLy$()) As String()
FmtAyDot1 = FmtDryAsSpcSep(DryzDotLyzTwoCol(DotLy))
End Function

Function FmtSyT1(Sy$()) As String()
FmtSyT1 = FmtAyNTerm(Sy, 1)
End Function

Function FmtSyT2(Sy$()) As String()
FmtSyT2 = FmtAyNTerm(Sy, 2)
End Function

Function FmtSyT3(Sy$()) As String()
FmtSyT3 = FmtAyNTerm(Sy, 3)
End Function

Function FmtSyT4(Sy$()) As String()
FmtSyT4 = FmtAyNTerm(Sy, 4)
End Function


Private Sub Z_FmtSyT2()
Dim Ly$()
Ly = Sy("AAA B C D", "Sy BBB CCC")
Ept = Sy("AAA B   C D", _
         "Sy   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtSyT2(Ly)
    C
    Return
End Sub
Private Sub Z_FmtSyT3()
Dim Ly$()
Ly = Sy("AAA B C D", "Sy BBB CCC")
Ept = Sy("AAA B   C   D", _
         "Sy   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtSyT3(Ly)
    C
    Return
End Sub

Private Sub Z()
Z_FmtSyT2
Z_FmtSyT3
MVb_Align_Ay:
End Sub
