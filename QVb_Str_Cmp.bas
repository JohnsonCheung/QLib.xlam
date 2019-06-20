Attribute VB_Name = "QVb_Str_Cmp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Str_Cmp."

Sub CmprLines(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$)
Brw FmtCmprLines(A, B, N1, N2, Hdr)
End Sub

Function FmtCmprLines(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
ThwIf_NotStr A, CSub
ThwIf_NotStr B, CSub
If A = B Then Exit Function
Dim AA$(), BB$()
AA = SplitCrLf(A)
BB = SplitCrLf(B)
If IsEqAy(AA, BB) Then Exit Function
Dim DifAt&
    DifAt = DifAtIx(AA, BB)
Dim O$(), J&, MinU&
    PushNB O, Hdr
    PushI O, FmtQQ("LinesCnt=? (?)", Si(AA), N1)
    PushI O, FmtQQ("LinesCnt=? (?)", Si(BB), N2)
    PushI O, FmtQQ("Dif At Ix=?", DifAt)
    
    MinU = Min(UB(AA), UB(BB))
    For J = 0 To MinU
        PushIAy O, LyzCmprStr(AA(J), BB(J), J)
    Next
    PushIAy O, LyRest(AA, BB, MinU, N1, N2)
    PushIAy O, LyAll(AA, N1)
    PushIAy O, LyAll(BB, N2)
FmtCmprLines = O
End Function
Private Function LyAll(A$(), Nm$) As String()
Stop
End Function
Private Function LyzCmprStr(A$, B$, Ix) As String()
If A = B Then PushI LyzCmprStr, Ix & ":" & A: Exit Function
PushI LyzCmprStr, Ix & ":" & A & "<" & Len(A)
Dim W%
W = Len(CStr(Ix)) + 1
PushI LyzCmprStr, Space(W) & B & "<" & Len(B)
End Function
Private Function LyRest(A$(), B$(), MinU&, Nm1$, Nm2$) As String()
Dim Ay$(), Nm$
Select Case True
Case UB(A) > MinU: Ay = A: Nm = Nm1
Case UB(B) > MinU: Ay = B: Nm = Nm2
Case Else: Exit Function
End Select
PushI LyRest, FmtQQ("Rest of (?) ------------", Nm)
Dim Pfx$, J&
For J = MinU + 1 To UB(Ay)
    Pfx = J & ":"
    PushI LyRest, J & ":" & Ay(J)
Next
End Function

Sub CmprStr(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$, Optional WiHdr As EmHdr = EmHdr.EiWiHdr)
If A = B Then Exit Sub
Brw FmtCmprStr(A, B, N1, N2, Hdr, WiHdr)
End Sub

Function FmtCmprStr(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$, Optional WiHdr As EmHdr = EmHdr.EiWiHdr) As String()
If Not IsStr(A) Then Stop
If Not IsStr(B) Then Stop
If IsLines(A) Or IsLines(B) Then FmtCmprStr = FmtCmprLines(A, B, N1, N2, Hdr): Exit Function
If A = B Then Exit Function
Dim At&
    At = DifAt(A, B)
Dim O$()
    If WiHdr = EiWiHdr Then
        PushI O, FmtQQ("Str-(?)-Len: ?", N1, Len(A))
        PushI O, FmtQQ("Str-(?)-Len: ?", N2, Len(B))
        PushI O, "Dif At: " & At
        PushIAy O, Len_LblAy(Max(Len(A), Len(B)))
    End If
    PushI O, A
    PushI O, B
    PushI O, Space(At - 1) & "^"
FmtCmprStr = O
End Function

Private Function DifAt&(A$, B$)
Dim O&
For O = 1 To Min(Len(A), Len(B))
    If Mid(A, O, 1) <> Mid(B, O, 1) Then DifAt = O: Exit Function
Next
If Len(A) > Len(B) Then
    DifAt = Len(B) + 1
Else
    DifAt = Len(A) + 1
End If
End Function

Private Function DifAtIx&(A$(), B$())
Dim O&
For O = 0 To Min(UB(A), UB(B))
    If A(O) <> B(O) Then DifAtIx = O: Exit Function
Next
'Thw_Never CSub
End Function

Function Len_LblAy(L&) As String()
Const CSub$ = CMod & "Len_LblAy"
If L <= 0 Then Thw CSub, "Length should be >0", "Length", L
Dim N%
    N = NDig(L)
PushNB Len_LblAy, Len_LblLin1(L)
PushI Len_LblAy, Len_LblLin2(L)
End Function

Private Function Len_LblLin1$(L&)
Dim J&, O$(), N&
PushI O, Space(9)
For J = 1 To (L - 1) \ 10 + 1
    N = J * 10
    PushI O, N & Space(10 - NDig(N))
Next
Len_LblLin1 = Join(O, "")
End Function

Private Function Len_LblLin2$(L&)
Dim Q&, R%
Const C$ = "123456789 "
Q = (L - 1) \ 10 + 1
R = (L - 1) Mod 10 + 1
Len_LblLin2 = Dup(C, Q) & Left(C, R)
End Function

Private Sub Z_FmtCmprLines()
Dim A$, B$
A = LineszVbl("AAAAAAA|bbbbbbbb|cc|dd")
B = LineszVbl("AAAAAAA|bbbbbbbb |cc")
GoSub Tst
Exit Sub
Tst:
    Act = FmtCmprLines(A, B)
    Brw Act
    Return

End Sub

Private Sub Z()
End Sub

