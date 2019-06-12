Attribute VB_Name = "QVb_Str_Tak"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Tak."
Private Const Asm$ = "QVb"

Function BefDot$(S)
BefDot = Bef(S, ".")
End Function
Function BefDotOrAll$(S)
BefDotOrAll = BefOrAll(S, ".")
End Function

Function Aft$(S, Sep$, Optional NoTrim As Boolean)
Aft = Brk1(S, Sep, NoTrim).S2
End Function

Function AftAt$(S, At&, Sep$)
If At = 0 Then Exit Function
AftAt = Mid(S, At + Len(Sep))
End Function

Function AftDotOrAll$(S)
AftDotOrAll = AftOrAll(S, ".")
End Function

Function AftDot$(S)
AftDot = Aft(S, ".")
End Function

Function AftMust$(S, Sep$, Optional NoTrim As Boolean)
AftMust = Brk(S, Sep, NoTrim).S2
End Function

Function AftOrAll$(S, Sep$, Optional NoTrim As Boolean)
AftOrAll = Brk2(S, Sep, NoTrim).S2
End Function

Function AftOrAllRev$(S, Sep$)
AftOrAllRev = StrDft(AftRev(S, Sep), Sep)
End Function

Function AftRev$(S, Sep$, Optional NoTrim As Boolean)
AftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function
Function BefSpc$(S)
BefSpc = Bef(S, " ")
End Function
Function AftSpc$(S)
AftSpc = Aft(S, " ")
End Function
Function BefSpcOrAll$(S)
BefSpcOrAll = BefOrAll(S, " ")
End Function
Function BefSy(Sy$(), Sep$, Optional NoTrim As Boolean) As String()
Dim I
For Each I In Itr(Sy)
    PushI BefSy, Bef(I, Sep, NoTrim)
Next
End Function
Function Bef$(S, Sep$, Optional NoTrim As Boolean)
Bef = Brk2(S, Sep, NoTrim).S1
End Function

Function RmvBef$(S, Sep$, Optional NoTrim As Boolean)
RmvBef = Brk2(S, Sep, NoTrim).S2
End Function

Function BefAt(S, At&)
If At = 0 Then Exit Function
BefAt = Left(S, At - 1)
End Function

Function BefDD$(S)
BefDD = RTrim(BefOrAll(S, "--"))
End Function

Function BefDDD$(S)
BefDDD = RTrim(BefOrAll(S, "---"))
End Function

Function BefMust$(S, Sep$, Optional NoTrim As Boolean)
BefMust = Brk(S, Sep, NoTrim).S1
End Function

Function BefOrAll$(S, Sep$, Optional NoTrim As Boolean)
BefOrAll = Brk1(S, Sep, NoTrim).S1
End Function

Function BefOrAllRev$(S, Sep$)
BefOrAllRev = StrDft(BefRev(S, Sep), Sep$)
End Function

Function BefRev$(S, Sep$, Optional NoTrim As Boolean)
BefRev = Brk2Rev(S, Sep, NoTrim).S1
End Function
Function P123(S, S1$, S2$) As String()
Dim P1&, P2&
P1 = InStr(S, S1)
P2 = InStr(P1 + Len(S1), S, S2)
If P2 > P1 And P1 > 0 And P2 > 0 Then
    PushI P123, Left(S, P1)
    Dim L&
        L = P2 - P1 - Len(S1)
    PushI P123, Mid(S, P1 + Len(S1), L)
    PushI P123, Mid(S, P2 + Len(S2))
End If
End Function
Sub AsgP123(S, S1$, S2$, O1$, O2$, O3$)
AsgAp P123(S, S1, S2), O1, O2, O3
End Sub
Private Sub Z_Tak_BefFstLas()
Dim S, Fst$, Las$
S = " A_1$ = ""Private Function ZChunk$(ConstLy$(), IChunk%)"" & _"
Fst = vbDblQuote
Las = vbDblQuote
Ept = "Private Function ZChunk$(ConstLy$(), IChunk%)"
GoSub Tst
Exit Sub
Tst:
    Act = BetFstLas(S, Fst, Las)
    C
    Return
End Sub
Function BetFstLas$(S, Fst$, Las$)
BetFstLas = BefRev(Aft(S, Fst), Las)
End Function
Function BetLng(L&, A&, B&) As Boolean
BetLng = A <= L And L <= B
End Function

Function Bet$(S, S1$, S2$, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   Bet = O
End With
End Function

Private Sub Z_BetBkt()
Dim Act$
   Dim S
   S = "sdklfjdsf(1234()567)aaa("
   Act = BetBkt(S)
   Ass Act = "1234()567"
End Sub
Function Nm$(S)
Nm = TakNm(S)
End Function
Function TakDotNm$(S)
Dim J%
If Not IsLetter(FstChr(S)) Then Exit Function
For J = 2 To Len(S)
    If Not IsDotNmChr(Mid(S, J, 1)) Then
        TakDotNm = Left(S, J - 1)
        Exit Function
    End If
Next
TakDotNm = S
End Function
Function TakNmzSy(Sy$()) As String()
Dim S
For Each S In Itr(Sy)
    PushI TakNmzSy, TakNm(S)
Next
End Function
Function TakNm$(S)
Dim J%
If Not IsLetter(FstChr(S)) Then Exit Function
For J = 2 To Len(S)
    If Not IsNmChr(Mid(S, J, 1)) Then
        TakNm = Left(S, J - 1)
        Exit Function
    End If
Next
TakNm = S
End Function

Function Pfx$(S, P$) ' Return Pfx-P if S has such Pfx else return ""
If HasPfx(S, P) Then Pfx = P
End Function

Function PfxSySpc$(S, PfxSy$()) ' Return Fst ele-P of [PfxSy] if [S] has pfx ele-P and a space
Dim P
For Each P In PfxSy
    If HasPfx(S, P & " ") Then PfxSySpc = P: Exit Function
Next
End Function

Function PfxzAy$(S, PfxSy$()) ' Return Fst ele-P of [PfxSy] if [S] has pfx ele-P
Dim P
For Each P In PfxSy
    If HasPfx(S, P) Then PfxzAy = P: Exit Function
Next
End Function

Function SfxzAySpc$(S, SfxSy$()) ' Return Fst ele-P of [PfxSy] if [S] has pfx ele-P
Dim Sfx$, I
For Each I In SfxSy
    Sfx = I
    If HasSfx(S, Sfx) Then SfxzAySpc = S: Exit Function
Next
End Function

Function PfxS(S, Pfx$) ' Return [Pfx] if [S] has such pfx+" " else return ""
If HasPfx(S, Pfx) Then If Mid(S, Len(Pfx) + 1, 1) = " " Then PfxS = Pfx
End Function

Private Sub Z_AftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = AftBkt(A)
    C
    Return
End Sub

Private Sub Z_Bet()
Dim Lin
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Lin = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Lin = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AsgAp AyTrim(SplitVBar(Lin)), Lin, FmStr, ToStr, Ept
    Act = Bet(Lin, FmStr, ToStr)
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
    Act = BetBkt(A)
    C
    Return
End Sub

Private Sub ZZ()
Z_AftBkt
Z_Tak_BefFstLas
Z_Bet
MVb_Str_Tak:
End Sub

Function BefRevOrAll$(S, Sep$)
Dim P%: P = InStrRev(S, Sep)
If P = 0 Then BefRevOrAll = S: Exit Function
BefRevOrAll = Left(S, P - Len(Sep))
End Function
'
'Function AftRev$(S, Sep$)
'Dim P%: P = InStrRev(S, Sep): If P = 0 Then Exit Function
'AftRev = Mid(S, P + Len(Sep))
'End Function
'

