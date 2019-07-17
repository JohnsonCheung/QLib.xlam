Attribute VB_Name = "QVb_Str_Cmpr"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Str_Cmp."

Sub CmprLines(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$)
Brw FmtCmprLines(A, B, N1, N2, Hdr)
End Sub

Private Function FmtCmprLines(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$ = "Compare 2 Lines") As String()
If A = B Then Exit Function
Dim AA$(): AA = SplitCrLf(A)
Dim BB$(): BB = SplitCrLf(B)
Dim DifAt&: DifAt = DifAtIx(AA, BB)
Dim O$()
    PushNB O, Box(Hdr)
    PushI O, FmtQQ("LinesCnt=? (?)", Si(AA), N1)
    PushI O, FmtQQ("LinesCnt=? (?)", Si(BB), N2)
    PushI O, FmtQQ("Dif At Ix=?", DifAt)
    
    '-- Sam ---
    PushI O, FmtQQ("-- Sam (0-?)---------", DifAt - 1)
    PushIAy O, AddIxPfx(AwFstUEle(A, DifAt - 1))
    
    '-- Dif Lin ---
    PushI O, FmtQQ("-- Dif (?)---------", DifAt - 1)
    PushIAy O, FmtCmprStr(AA(DifAt), BB(DifAt), "", "")
    
    '-- Rst-A & B---
    PushIAy O, FmtCmprLines__Rst(AA, DifAt, N1)
    PushIAy O, FmtCmprLines__Rst(BB, DifAt, N2)
FmtCmprLines = O
End Function

Private Function FmtCmprLines__Rst(A$(), DifAt&, N1$) As String()
Dim O$()
PushI O, FmtQQ("-- Rst-? (?-?) ----------", N1, DifAt + 1, UB(A))
PushIAy O, AddIxPfx(AwFm(A, DifAt + 1), EiBegI, DifAt + 1)
FmtCmprLines__Rst = O
End Function
Sub CmprStr(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$)
If A = B Then Exit Sub
Brw FmtCmprStr(A, B, N1, N2, Hdr)
End Sub

Function FmtCmprStr(A$, B$, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
'== Case1 A=B ===
If A = B Then
    PushIAy FmtCmprStr, Box(Hdr)
    PushI FmtCmprStr, FmtQQ("Str(?) = Str(?).  Len(?)", N1, N2, Len(A))
    Exit Function
End If
'== Case2 IsLines ==
Select Case True
Case IsLines(A), IsLines(B)
    FmtCmprStr = FmtCmprLines(A, B, N1, N2, Hdr)
    Exit Function
End Select
'== Case3 IsStr
Dim At&: At = DifAt(A, B)
Dim O$()
    PushI O, FmtQQ("Str-(?)-Len: ?", N1, Len(A))
    PushI O, FmtQQ("Str-(?)-Len: ?", N2, Len(B))
    PushI O, "Dif At: " & At
    PushIAy O, Lbl123(Max(Len(A), Len(B)))
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
Private Sub Z_Lbl123()
Dmp Lbl123(543)
End Sub
Function Lbl123(L) As String()
'Fm L : 1 to 999, else Thw
'Ret  : :Lbl123: is sy-of-1-to-3 ele.  Las-ele is dig no 0, Las-2nd-if-any is ten-dig, Las-3rd-if-any is hundred-dig
Const CSub$ = CMod & "Lbl12"
If Not IsBet(L, 1, 999) Then Thw CSub, "Length should be bet 1 999", "Length", L
PushNB Lbl123, Lbl123__Hundred(L)
PushNB Lbl123, Lbl123__Ten(L)
PushI Lbl123, Lbl123__Dig(L)
End Function

Private Function Lbl123__Dig$(L)
Const C$ = "1234567890"
Dim N&: N = (L \ 10) + 1
Lbl123__Dig = Left(Dup(C, N), L)
End Function

Private Function Lbl123__Ten$(L)
If L < 9 Then Exit Function
Dim O$()
    PushI O, Space(9)
    Dim J%: For J = 0 To (L \ 10)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, Dup(C, 10)
    Next
Lbl123__Ten = Left(Jn(O), L)
End Function

Private Function Lbl123__Hundred$(L)
If L < 99 Then Exit Function
Dim O$()
    PushI O, Space(99)
    Dim J%: For J = 0 To (L \ 100)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, Dup(C, 100)
    Next
Lbl123__Hundred = Left(Jn(O), L)
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
