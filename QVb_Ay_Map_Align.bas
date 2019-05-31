Attribute VB_Name = "QVb_Ay_Map_Align"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Map_Align."
Private Const Asm$ = "QVb"
Enum EmAlign
    EiLeft
    EiRight
End Enum
Function FmtSyzNTerm(Ay, N%) As String()
Dim W%(), L
W = WdtAyNTermAy(N, Ay)
For Each L In Itr(Ay)
    PushI FmtSyzNTerm, FmtSyzNTerm1(L, W)
Next
End Function

Private Function FmtSyzNTerm1$(Sy, W%())
Dim Ay$(), J%, N%, O$(), I
N = Si(W)
Ay = SyzNTermRst(Sy, N)
If Si(Ay) <> N + 1 Then Stop
For J = 0 To N - 1
    PushI O, AlignL(Ay(J), W(J))
Next
PushI O, Ay(N)
FmtSyzNTerm1 = RTrim(JnSpc(O))
End Function

Private Function WdtAyNTermAy(NTerm%, Ay) As Integer()
If Si(Ay) = 0 Then Exit Function
Dim O%(), W%(), L
ReDim O(NTerm - 1)
For Each L In Ay
    W = WdtAyNTermLin(NTerm, L)
    O = WdtAyab(O, W)
Next
WdtAyNTermAy = O
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
Function S1S2zAtChr(S, AtChr$, Optional IfNoAtChr As EmAlign) As S1S2
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

Function FmtSyzAtChr(Sy$(), AtChr$, Optional IfNoAtChr As EmAlign) As String()
FmtSyzAtChr = FmtS1S2s(S1S2szSyAtChr(Sy, AtChr))
End Function

Function FmtSyzAtDot(Sy$(), Optional IfNoDt As EmAlign) As String()
FmtSyzAtDot = FmtSyzAtChr(Sy, ".")
End Function

Sub BrwDotLy(DotLy$())
Brw FmtDotLy(DotLy)
End Sub

Function FmtDotLy(DotLy$()) As String()
FmtDotLy = FmtDry(DryzDotLy(DotLy), Fmt:=EiSSFmt)
End Function

Function FmtDotLyzTwoCol(DotLy$()) As String()
FmtDotLyzTwoCol = FmtDry(DryzDotLyzTwoCol(DotLy), Fmt:=EiSSFmt)
End Function

Function FmtSyz1Term(Sy$()) As String()
FmtSyz1Term = FmtSyzNTerm(Sy, 1)
End Function

Function FmtSyz2Term(Sy$()) As String()
FmtSyz2Term = FmtSyzNTerm(Sy, 2)
End Function

Function FmtSy3Term(Sy$()) As String()
FmtSy3Term = FmtSyzNTerm(Sy, 3)
End Function

Function FmtSyT4(Sy$()) As String()
FmtSyT4 = FmtSyzNTerm(Sy, 4)
End Function


Private Sub Z_FmtSyz2Term()
Dim Ly$()
Ly = Sy("AAA B C D", "Sy BBB CCC")
Ept = Sy("AAA B   C D", _
         "Sy   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtSyz2Term(Ly)
    C
    Return
End Sub
Private Sub Z_FmtSy3Term()
Dim Ly$()
Ly = Sy("AAA B C D", "Sy BBB CCC")
Ept = Sy("AAA B   C   D", _
         "Sy   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtSy3Term(Ly)
    C
    Return
End Sub

Private Sub ZZ()
Z_FmtSyz2Term
Z_FmtSy3Term
MVb_Align_Ay:
End Sub
