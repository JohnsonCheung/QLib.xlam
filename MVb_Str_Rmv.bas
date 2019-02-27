Attribute VB_Name = "MVb_Str_Rmv"
Option Explicit

Function RmvDotComma$(A)
RmvDotComma = Replace(Replace(A, ",", ""), ".", "")
End Function
Function Rmv2Dash$(A)
Rmv2Dash = RTrim(RmvAft(A, "--"))
End Function

Function Rmv3Dash$(A)
Rmv3Dash = RTrim(RmvAft(A, "---"))
End Function

Function Rmv3T$(A$)
Rmv3T = RmvTT(RmvT1(A))
End Function

Function RmvAft$(A, Sep$)
RmvAft = Brk1(A, Sep, NoTrim:=True).S1
End Function

Function RmvDDRmk$(A$)
Dim S$
If HasDDRmk(A) Then
    S = ""
Else
    S = A
End If
End Function

Function RmzlSpc$(A)
Dim O$: O = A
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmzlSpc = O
End Function

Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Function

Function RmvFstNChr$(A, Optional N% = 1)
RmvFstNChr = Mid(A, N + 1)
End Function

Function RmvFstNonLetter$(A)
If IsAscLetter(Asc(A)) Then
    RmvFstNonLetter = A
Else
    RmvFstNonLetter = RmvFstChr(A)
End If
End Function

Function RmvLasChr$(A)
RmvLasChr = RmvLasNChr(A, 1)
End Function

Function RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - N)
End Function

Function RmvNm$(A)
Dim O%
If Not IsAscFstNmChr(Asc(FstChr(A))) Then GoTo X
For O = 1 To Len(A)
    If Not IsAscNmChr(Asc(Mid(A, O, 1))) Then GoTo X
Next
X:
    If O > 0 Then RmvNm = Mid(A, O): Exit Function
    RmvNm = A
End Function

Function RmvSqBkt$(A)
If Not HasSqBkt(A) Then RmvSqBkt = A: Exit Function
RmvSqBkt = RmvFstLasChr(A)
End Function

Function RmvPfx$(A, Pfx)
If HasPfx(A, Pfx) Then RmvPfx = Mid(A, Len(Pfx) + 1) Else RmvPfx = A
End Function

Function RmvPfxAy$(A, PfxAy)
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then RmvPfxAy = RmvPfx(A, Pfx): Exit Function
Next
RmvPfxAy = A
End Function
Function RmvPfxSpc$(A, Pfx)
If Not HitPfxSpc(A, Pfx) Then RmvPfxSpc = A: Exit Function
RmvPfxSpc = LTrim(Mid(A, Len(Pfx) + 2))
End Function
Function RmvPfxAySpc$(A, PfxAy)
Dim P
For Each P In PfxAy
    If HitPfxSpc(A, P) Then
        RmvPfxAySpc = LTrim(Mid(A, Len(P) + 2))
        Exit Function
    End If
Next
RmvPfxAySpc = A
End Function

Function RmvSfx$(A, Sfx)
If HasSfx(A, Sfx) Then RmvSfx = Left(A, Len(A) - Len(Sfx)) Else RmvSfx = A
End Function

Function RmvSngQuote$(A)
If Not IsSngQuoted(A) Then RmvSngQuote = A: Exit Function
RmvSngQuote = RmvFstLasChr(A)
End Function

Function RmvT1$(A)
RmvT1 = SyTRst(A)(1)
End Function

Function RmvTT$(A)
RmvTT = RmvT1(RmvT1(A))
End Function

Private Sub Z_RmvT1()
Ass RmvT1("  df dfdf  ") = "dfdf"
End Sub


Private Sub Z_RmvNm()
Dim Nm$
Nm = "lksdjfsd f"
Ept = " f"
GoSub Tst
Exit Sub
Tst:
    Act = RmvNm(Nm)
    C
    Return
End Sub

Private Sub Z_RmvPfx()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub

Private Sub Z_RmvPfxAy()
Dim A$, PfxAy$()
PfxAy = SySsl("ZZ_ Z_"): Ept = "ABC"
A = "Z_ABC": GoSub Tst
A = "ZZ_ABC": GoSub Tst
Exit Sub
Tst:
    Act = RmvPfxAy(A, PfxAy)
    C
    Return
End Sub


Private Sub Z()
MVb_Str_Rmv:
End Sub
