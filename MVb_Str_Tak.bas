Attribute VB_Name = "MVb_Str_Tak"
Option Explicit

Function StrBefDot$(A)
StrBefDot = StrBef(A, ".")
End Function

Function StrAft$(S, Sep, Optional NoTrim As Boolean)
StrAft = Brk1(S, Sep, NoTrim).S2
End Function

Function StrAftAt$(A, At&, S)
If At = 0 Then Exit Function
StrAftAt = Mid(A, At + Len(S))
End Function

Function StrAftDotOrAll$(A)
StrAftDotOrAll = StrAftOrAll(A, ".")
End Function

Function StrAftDot$(A)
StrAftDot = StrAft(A, ".")
End Function

Function StrAftMust$(A, Sep, Optional NoTrim As Boolean)
StrAftMust = Brk(A, Sep, NoTrim).S2
End Function

Function StrAftOrAll$(S, Sep, Optional NoTrim As Boolean)
StrAftOrAll = Brk2(S, Sep, NoTrim).S2
End Function

Function StrAftOrAllRev$(A, S)
StrAftOrAllRev = StrDft(StrAftRev(A, S), A)
End Function

Function StrAftRev$(S, Sep, Optional NoTrim As Boolean)
StrAftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function StrBef$(S, Sep, Optional NoTrim As Boolean)
StrBef = Brk2(S, Sep, NoTrim).S1
End Function

Function StrBefAt(A, At&)
If At = 0 Then Exit Function
StrBefAt = Left(A, At - 1)
End Function

Function StrBefDD$(A)
StrBefDD = RTrim(StrBefOrAll(A, "--"))
End Function

Function StrBefDDD$(A)
StrBefDDD = RTrim(StrBefOrAll(A, "---"))
End Function

Function StrBefMust$(S, Sep$, Optional NoTrim As Boolean)
StrBefMust = Brk(S, Sep, NoTrim).S1
End Function

Function StrBefOrAll$(S, Sep, Optional NoTrim As Boolean)
StrBefOrAll = Brk1(S, Sep, NoTrim).S1
End Function

Function StrBefOrAllRev$(A, S)
StrBefOrAllRev = StrDft(StrBefRev(A, S), A)
End Function

Function StrBefRev$(A, Sep, Optional NoTrim As Boolean)
StrBefRev = Brk2Rev(A, Sep, NoTrim).S1
End Function
Function TakP123(A, S1, S2) As String()
Dim P1&, P2&
P1 = InStr(A, S1)
P2 = InStr(P1 + Len(S1), A, S2)
If P2 > P1 And P1 > 0 And P2 > 0 Then
    PushI TakP123, Left(A, P1)
    Dim L&
        L = P2 - P1 - Len(S1)
    PushI TakP123, Mid(A, P1 + Len(S1), L)
    PushI TakP123, Mid(A, P2 + Len(S2))
End If
End Function
Sub TakP123Asg(A, S1, S2, O1, O2, O3)
AsgAp TakP123(A, S1, S2), O1, O2, O3
End Sub
Private Sub Z_Tak_BefFstLas()
Dim S, Fst, Las
S = " A_1$ = ""Private Function ZChunk$(ConstLy$(), IChunk%)"" & _"
Fst = vbDblQuote
Las = vbDblQuote
Ept = "Private Function ZChunk$(ConstLy$(), IChunk%)"
GoSub Tst
Exit Sub
Tst:
    Act = StrBetFstLas(S, Fst, Las)
    C
    Return
End Sub
Function StrBetFstLas$(S, Fst, Las)
StrBetFstLas = StrBefRev(StrAft(S, Fst), Las)
End Function

Function StrBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   StrBet = O
End With
End Function

Private Sub Z_Tak_BetBkt()
Dim Act$
   Dim S$
   S = "sdklfjdsf(1234()567)aaa("
   Act = StrBetBkt(S)
   Ass Act = "1234()567"
End Sub

Function TakNm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        TakNm = Left(A, J - 1)
        Exit Function
    End If
Next
TakNm = A
End Function

Function TakPfx$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
If HasPfx(Lin, Pfx) Then TakPfx = Pfx
End Function

Function PfxAyFstSpc$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
Dim P
For Each P In PfxAy
    If HasPfx(Lin, P & " ") Then PfxAyFstSpc = P: Exit Function
Next
End Function

Function PfxLinAy$(A, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim P
For Each P In PfxAy
    If HasPfx(A, P) Then PfxLinAy = P: Exit Function
Next
End Function

Function SfxLinAy$(A, SfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim S
For Each S In SfxAy
    If HasSfx(A, S) Then SfxLinAy = S: Exit Function
Next
End Function

Function TermLinAy$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
TermLinAy = PfxAyFstSpc$(PfxAy, Lin)
End Function

Function TakPfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
If HasPfx(Lin, Pfx) Then If Mid(Lin, Len(Pfx) + 1, 1) = " " Then TakPfxS = Pfx
End Function

Function TakT1$(A)
If FstChr(A) <> "[" Then TakT1 = StrBef(A, " "): Exit Function
Dim P%
P = InStr(A, "]")
If P = 0 Then Stop
TakT1 = Mid(A, 2, P - 2)
End Function

Private Sub Z_StrAftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = StrAftBkt(A)
    C
    Return
End Sub

Private Sub Z_StrBet()
Dim Lin$
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Lin = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AsgAp AyTrim(SplitVBar(Lin)), Lin, FmStr, ToStr, Ept
    Act = StrBet(Lin, FmStr, ToStr)
    C
    Return
End Sub

Private Sub ZZ_Tak_BetBkt()
Dim A$
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Ept = "AA":        A = "XXX(AA)XX":   GoSub Tst
Ept = "A$()A":     A = "(A$()A)XX":   GoSub Tst
Ept = "O$()":      A = "(O$()) As X": GoSub Tst
Exit Sub
Tst:
    Act = StrBetBkt(A)
    C
    Return
End Sub

Private Sub Z()
Z_StrAftBkt
Z_Tak_BefFstLas
Z_StrBet
Z_Tak_BetBkt
MVb_Str_Tak:
End Sub

Function StrBefRevOrAll$(S, Sep$)
Dim P%: P = InStrRev(S, Sep)
If P = 0 Then StrBefRevOrAll = S: Exit Function
StrBefRevOrAll = Left(S, P - Len(Sep))
End Function
'
'Function StrAftRev$(S, Sep$)
'Dim P%: P = InStrRev(S, Sep): If P = 0 Then Exit Function
'StrAftRev = Mid(S, P + Len(Sep))
'End Function
'

