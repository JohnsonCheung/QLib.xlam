Attribute VB_Name = "MVb_Lin_Lines"
Option Explicit
Const CMod$ = "MVb_Lin_Lines."
Function LinesCntSzStr$(Lines$)
LinesCntSzStr = FmtQQ("LinesCntSz(? ?)", LinCnt(Lines), Len(Lines))
End Function
Private Sub Z_LinesWrap()
Dim A$, W%
A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
W = 80
Ept = Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
"klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
"sdklf sdklfj dsfj ")
GoSub Tst
Exit Sub
Tst:
    Act = LinesWrap(A, W)
    C
    Return
End Sub
Function LyWrap(Lines, Optional Wdt% = 80) As String()
LyWrap = SplitCrLf(LinesWrap(Lines, Wdt))
End Function

Function LinesWrap$(Lines, Optional Wdt% = 80)
Const CSub$ = CMod & "LinesWrap"
Dim L$, W%, O$, J%
W = Wdt
If W < 10 Then W = 10: Info CSub, "Given Wdt is too small, 10 is used", "Wdt Lines", Wdt, Lines
O = Lines
While Len(L) > 0
    J = J + 1: If J >= 1000 Then Stop
    O = LinesWrJnSpcAp(O, W)
    L = LinesWrapRst(O, W)
    PushI LinesWrap, L
Wend
End Function
Private Function LinesWrJnSpcAp$(A$, W%)
LinesWrJnSpcAp = Left(A, W)
End Function

Private Function LinesWrapRst$(A$, W%)
LinesWrapRst = Mid(A, W + 1)
Exit Function
Dim O$, L$
    If FstChr(O) = " " Then
        O = LTrim(O)
    Else
        If LasChr(L) <> " " Then
            Dim P%
            P = InStrRev(L, " ")
            If P = 0 Then
                O = ""
            Else
                O = Mid(L, P + 1) & O
                L = Left(L, P - 1)
            End If
        End If
    End If

End Function

Private Sub ZZ_TrimRCrLf()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = TrimRCrLf(Lines)
Debug.Print Act & "<"
Stop
End Sub

Private Sub ZZ_LasNLines()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
Debug.Print LasNLines(A, 3)
End Sub
Function FstLin$(Lines)
FstLin = RplCr(TakBefOrAll(Lines, vbLf))
End Function
Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Function SplitCrLfAy(LinesAy) As String()
Dim Lines
For Each Lines In Itr(LinesAy)
    PushIAy SplitCrLfAy, SplitCrLf(Lines)
Next
End Function

Sub LinesAsgBrk(A$, Ny0, ParamArray OLyAp())
Dim Ny$(), L, T1$, T2$, NmDic As Dictionary
Ny = CvNy(Ny0)
Set NmDic = IxDiczAy(Ny)
For Each L In SplitCrLf(A)
    Select Case FstChr(L)
    Case "'", " "
    Case Else
        AsgBrk L, " ", T1, T2
        If NmDic.Exists(T1) Then
            Push OLyAp(NmDic(T1)), T2 '<----
        End If
    End Select
Next
End Sub

Private Sub Z_TrimRCrLf()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = TrimRCrLf(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LasNLines$(Lines, N%)
LasNLines = JnCrLf(AywLasN(SplitCrLf(Lines), N))
End Function

Function LinCnt&(Lines)
LinCnt = Sz(SplitCrLf(Lines))
End Function

Function HSqLines(Lines) As Variant()
HSqLines = SqzAyH(SplitCrLf(Lines))
End Function

Function VSqLines(Lines) As Variant()
VSqLines = SqzAyV(SplitCrLf(Lines))
End Function

Function TrimR$(S)
TrimR = TrimRCrLf(RTrim(S))
End Function

Function RLenOfCrLf%(S)
Dim O%, J%
For J = Len(S) To 1 Step -1
    If IsCrLf(AscAt(S, J)) Then O = O + 1
Next
End Function

Function AscAt%(S, Pos)
AscAt = Asc(Mid(S, Pos, 1))
End Function

Function IsCrLf(Asc%)
IsCrLf = Asc = 13 Or Asc = 10
End Function
Function TrimRCrLf$(S)
TrimRCrLf = RmvLasNChr(S, RLenOfCrLf(S))
End Function

Function LasLinLines$(Lines)
LasLinLines = LasEle(SplitCrLf(Lines))
End Function
Function LinesAlignL$(Lines, W%)
Dim Las$: Las = LasLinLines(Lines)
Dim N%: N = W - Len(Las)
If N > 0 Then
    LinesAlignL = Lines & Space(N)
Else
    Warn CSub, "W is too small", "Lines.LasLin W", Las, W
    LinesAlignL = Lines
End If
End Function

Function NLines&(Lines$)
NLines = SubStrCnt(Lines, vbLf) + 1
End Function

