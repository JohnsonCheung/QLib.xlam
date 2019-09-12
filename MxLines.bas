Attribute VB_Name = "MxLines"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLines."
Function WdtzLines%(Lines)
WdtzLines = AyWdt(SplitCrLf(Lines))
End Function

Function WdtzLinesAy%(LinesAy)
Dim O%, Lines
For Each Lines In Itr(LinesAy)
    O = Max(O, WdtzLines(Lines))
Next
WdtzLinesAy = O
End Function

Sub VcLinesAy(LinesAy)
Vc FmtLinesAy(LinesAy)
End Sub

Sub BrwLinesAy(LinesAy)
B FmtLinesAy(LinesAy)
End Sub

Function MrkLines$(Lines, Lno&)
Dim Ly$(): Ly = SplitCrLf(Lines)

End Function

Private Sub Z_FmtLinesAy()
Dim LinesAy
GoSub Z
Exit Sub
Z:
    BrwLinesAy Y_LinesAy
    Return
End Sub

Private Function Y_LinesAy() As String()
Erase XX
X RplVbl("sdklf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
X RplVbl("sdklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
X RplVbl("sdsdfklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
Y_LinesAy = XX
Erase XX
End Function

Function FmtLinesAy(LinesAy, Optional B As EmIxCol = EiBeg0) As String()
If Si(LinesAy) = 0 Then Exit Function
Dim W%: W = WdtzLinesAy(LinesAy)
Dim LinzSep: LinzSep = Qte(Dup("-", W + 2), "|")
Dim Lines
PushI FmtLinesAy, LinzSep
For Each Lines In Itr(LinesAy)
    PushIAy FmtLinesAy, AddIxPfxzLineszW(Lines, W, B)
    PushI FmtLinesAy, LinzSep
Next
End Function
Private Function AddIxPfxzLineszW(Lines, W%, Optional B As EmIxCol = EiBeg0) As String()
Dim L
For Each L In Itr(SplitCrLf(Lines))
    PushI AddIxPfxzLineszW, "| " & AlignL(L, W) & " |"
Next
End Function


Private Sub Z_WrpLines()
Dim A$, W%
A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
W = 80
Ept = Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
"klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
"sdklf sdklfj dsfj ")
GoSub Tst
Exit Sub
Tst:
    Act = WrpLines(A, W)
    C
    Return
End Sub
Function WrpLines$(Lines$, Optional Wdt% = 80)
WrpLines = Lines: Exit Function
WrpLines = JnCrLf(WrpLy(SplitCrLf(Lines), Wdt))
End Function

Private Sub Z_WrpLy()
Dim Ly$(), Wdt%
GoSub T1
Exit Sub
T1:
    Ly = Sy("a b c d")
    Wdt = 80
    Ept = Sy("a b c d")
    GoTo Tst
Tst:
    Act = WrpLy(Ly, Wdt)
    C
    Return
End Sub
Function EnsBet%(I%, A%, B%)
Select Case True
Case I < A: EnsBet = A
Case I > B: EnsBet = B
Case Else: EnsBet = I
End Select
End Function

Function WrpLy(Ly$(), Optional Wdt% = 80) As String()
Dim W%, Lin, I
W = EnsBet(Wdt, 10, 200)
For Each I In Itr(Ly)
    Lin = I
    PushIAy WrpLy, LyzWrpLin(Lin, W)
Next
End Function

Private Function ShfWrpgLin(OLin, W%, LasLinLasChr$)
If OLin = "" Then Exit Function
Dim O$, OL$, F$
O = Left(OLin, W)
OL = Mid(OLin, W + 1)
F = FstChr(OL)
Select Case True
Case OL = "" Or F = " "
Case LasLinLasChr = " ": OL = LTrim(OL)
Case Else:
    Dim P%: P = InStrRev(O, " ")
    If P <> 0 Then
        O = Left(O, P)
        OL = Mid(O, P + 1) & OL
    End If
End Select
ShfWrpgLin = Trim(O)
OLin = OL
End Function

Private Function LyzWrpLin(Lin, W%) As String()
If Len(Lin) > W Then LyzWrpLin = Sy(Lin): Exit Function
Dim L$: L = RTrim(Lin)
Dim J%
Dim LasLinLasChr$
While L <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    PushI LyzWrpLin, ShfWrpgLin(L, W, LasLinLasChr)
    LasLinLasChr = LasChr(L)
Wend
End Function

Private Sub Z_LineszRTrim()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = LineszRTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Private Sub Z_LineszLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
Debug.Print LineszLasN(A, 3)
End Sub

Function FstLin(Lines$)
FstLin = BefOrAll(Lines, vbCrLf)
End Function

Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Private Sub Z_LineszRTrim1()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = LineszRTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LineszLasN$(Lines$, N%)
LineszLasN = JnCrLf(AwLasN(SplitCrLf(Lines), N))
End Function

Function LinCnt&(Lines)
LinCnt = Si(SplitCrLf(Lines))
End Function

Function MaxLinCnt&(LinesAy$())
Dim O&, Lines: For Each Lines In Itr(LinesAy)
    O = Max(O, LinCnt(Lines))
Next
MaxLinCnt = O
End Function

Function SqhzLines(Lines$) As Variant()
SqhzLines = SqH(SplitCrLf(Lines))
End Function

Function SqvzLines(Lines$) As Variant()
SqvzLines = SqV(SplitCrLf(Lines))
End Function

Function LineszRTrim$(Lines)
Dim At&
For At = Len(Lines) To 1 Step -1
    If Not IsStrAtSpcCrLf(Lines, At) Then LineszRTrim = Left(Lines, At): Exit Function
Next
End Function

Function LasLin(Lines$)
LasLin = LasEle(SplitCrLf(Lines))
End Function

Function LineszAlign$(Lines$, W%)
Dim LAs$: LAs = LasLin(Lines)
Dim N%: N = W - Len(LAs)
If N > 0 Then
    LineszAlign = Lines & Space(N)
Else
    Warn CSub, "W is too small", "Lines.LasLin W", LAs, W
    LineszAlign = Lines
End If
End Function

Function NLines&(Lines)
Dim R As RegExp: Set R = Rx("\n", MultiLine:=True, IsGlobal:=True)
NLines = Mch(R, Lines).Count + 1
End Function

Function LineszV$(V)
LineszV = JnCrLf(FmtV(V))
End Function

Function ColRgAyzLinesDr(LinesDr$()) As Variant()
Dim Lines: For Each Lines In LinesDr
    PushI ColRgAyzLinesDr, SplitCrLf(Lines)
Next
End Function

Function SqzLinesDr(LinesDr$()) As Variant()
Dim NRow%: NRow = MaxLinCnt(LinesDr)
Dim NCol%: NCol = Si(LinesDr)
Dim ColAy(): ColAy = ColRgAyzLinesDr(LinesDr)
Dim O(): ReDim O(1 To NRow, 1 To NCol)
Dim Col$(), ICol%, S$: For ICol = 0 To NCol - 1
    Col = ColAy(ICol)
    Dim IRow%: For IRow = 0 To UB(Col)
        S = Col(IRow)
        O(IRow + 1, ICol + 1) = S
    Next
Next
SqzLinesDr = O
End Function

