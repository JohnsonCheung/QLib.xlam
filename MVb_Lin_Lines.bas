Attribute VB_Name = "MVb_Lin_Lines"
Option Explicit
Const CMod$ = "MVb_Lin_Lines."
Function CntSzStrzLines$(Lines)
CntSzStrzLines = CntSzStr(LinCnt(Lines), Len(Lines))
End Function
Function CntSzStr$(Cnt&, Si&)
CntSzStr = FmtQQ("CntSzStr(? ?)", Cnt, Si)
End Function
Private Sub Z_LinesWrp()
Dim A$, W%
A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
W = 80
Ept = Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
"klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
"sdklf sdklfj dsfj ")
GoSub Tst
Exit Sub
Tst:
    Act = LinesWrp(A, W)
    C
    Return
End Sub
Function LinesWrp$(Lines, Optional Wdt% = 80)
LinesWrp = Lines: Exit Function
LinesWrp = JnCrLf(LyWrp(SplitCrLf(Lines), Wdt))
End Function
Private Sub Z_LyWrp()
Dim Ly$(), Wdt%
GoSub T1
Exit Sub
T1:
    Ly = Sy("a b c d")
    Wdt = 80
    Ept = Sy("a b c d")
    GoTo Tst
Tst:
    Act = LyWrp(Ly, Wdt)
    C
    Return
End Sub

Function LyWrp(Ly$(), Optional Wdt% = 80) As String()
Const CSub$ = CMod & "LinesWrp"
Dim L$(), W%, J%, A$, Lin
W = Wdt
If W < 10 Then W = 10: Inf CSub, "Given Wdt is too small, 10 is used", "Wdt Ly", Wdt, Ly
L = Ly
For Each Lin In Itr(Ly)
    PushIAy LyWrp, LyzLinWrp(Lin, W)
Next
End Function

Private Function LyzLinWrp(Lin, W%) As String()
If Len(Lin) > W Then LyzLinWrp = Sy(Lin): Exit Function
Dim L$: L = RTrim(Lin)
Dim O$(), A$, LasA$, FstL$, NoMore As Boolean, J%
Do
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    A = RTrim(Left(L, W))
    L = Mid(L, W + 1)
    LasA = LasChr(A)
    FstL = FstChr(L)
    NoMore = L = ""
    Select Case True
    Case NoMore: PushI LyzLinWrp, RTrim(A): Exit Function
    Case LasA = " " And FstL = " ": PushI LyzLinWrp, RTrim(A): L = LTrim(L)
    Case LasA = " ":                PushI LyzLinWrp, RTrim(A)
    Case FstL = " ":                PushI LyzLinWrp, A: L = LTrim(L)
    Case Else:
        Dim P%: P = InStrRev(A, " ")
        If P = 0 Then PushI LyzLinWrp, A: Exit Function
        PushI LyzLinWrp, Left(A, P - 1)
        L = LTrim(Mid(A, P + 1)) & L
    End Select
Loop
End Function

Private Sub ZZ_TrimCrLfAtEnd()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = TrimCrLfAtEnd(Lines)
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
FstLin = StrBefOrAll(Lines, vbCrLf)
End Function

Function LinesRmvBlankLinAtEnd$(Lines)
Dim J%, O$
O = Lines
Do
    J = J + 1: If J = 1000 Then ThwLoopingTooMuch CSub
    If HasSfx(O, vbCrLf) Then
        O = RmvSfx(O, vbCrLf)
    Else
        LinesRmvBlankLinAtEnd = O
        Exit Function
    End If
Loop
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
Ny = NyzNN(Ny0)
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

Private Sub Z_TrimCrLfAtEnd()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = TrimCrLfAtEnd(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LasNLines$(Lines, N%)
LasNLines = JnCrLf(AywLasN(SplitCrLf(Lines), N))
End Function

Function LinCnt&(Lines)
LinCnt = Si(SplitCrLf(Lines))
End Function

Function HSqLines(Lines) As Variant()
HSqLines = SqzAyH(SplitCrLf(Lines))
End Function

Function VSqLines(Lines) As Variant()
VSqLines = SqzAyV(SplitCrLf(Lines))
End Function

Function TrimR$(S)
TrimR = TrimCrLfAtEnd(RTrim(S))
End Function

Function RLenOfCrLf%(S)
End Function

Function AscAt%(S, Pos)
AscAt = Asc(Mid(S, Pos, 1))
End Function

Function IsAscCrLf(Asc%)
IsAscCrLf = (Asc = 13) Or (Asc = 10)
End Function

Function TrimCrLfAtEnd$(S)
Dim J&
For J = Len(S) To 1 Step -1
    If Not IsAscCrLf(AscAt(S, J)) Then TrimCrLfAtEnd = Left(S, J): Exit Function
Next
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

