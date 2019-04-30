Attribute VB_Name = "MVb_S1S2_Fmt"
Option Explicit
Function FmtS1S2s(A As S1S2s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As String()
Dim mHasLines As Boolean: mHasLines = HasLines(A)
Dim mSepChr$:               mSepChr = IIf(mHasLines, "|", " ")
Dim mS1$():                     mS1 = Sy1zS1S2s(A)
Dim mS2$():                     mS2 = Sy2zS1S2s(A)
Dim mW1%:                       mW1 = WdtzLinesAy(AddStrEle(mS1, Nm1))
Dim mW2%:                       mW2 = WdtzLinesAy(AddStrEle(mS2, Nm2))
Dim mSep$:                     mSep = SepLin(IntAy(mW1, mW2), mSepChr)
Dim mTit$:                     mTit = LinzS1S2(S1S2(Nm1, Nm2), mW1, mW2)
Dim mHdrLy$():               mHdrLy = Sy(mSep, mTit, mSep)
Dim mMidLy$():               mMidLy = LyzS1S2s(A, mW1, mW2, mHasLines, mSep)
FmtS1S2s = AyAdd(mHdrLy, mMidLy)
End Function

Private Function LyzS1S2(A As S1S2, W1%, W2%) As String()
Dim Lines1$, Lines2$
    Lines1 = A.S1
    Lines2 = A.S2
Dim NLin%
    NLin = Max(LinCnt(Lines1), LinCnt(Lines2))
Dim Ly1$(), Ly2$()
    Ly1 = SplitCrLf(Lines1)
    Ly2 = SplitCrLf(Lines2)
    ReSumSiabMax Ly1, Ly2
    Ly1 = SyAlignL(Ly1, W1)
    Ly2 = SyAlignL(Ly2, W2)
Dim J%, O$()
For J = 0 To UB(Ly1)
    PushI O, "| " & Ly1(J) & " | " & Ly2(J) & " |"
Next
LyzS1S2 = O
End Function

Private Function LyzS1S2s(A As S1S2s, W1%, W2%, S1S2sHasLines As Boolean, SepLin$) As String()
If S1S2sHasLines Then
    Dim J&
    For J = 0 To A.N - 1
        PushIAy LyzS1S2s, LyzS1S2(A.Ay(J), W1, W2)
        PushI LyzS1S2s, SepLin
    Next
Else
    For J = 0 To A.N - 1
        PushI LyzS1S2s, LinzS1S2(A.Ay(J), W1, W2)
    Next
    PushI LyzS1S2s, SepLin
End If
End Function
Private Function LinzS1S2$(A As S1S2, W1%, W2%)
LinzS1S2 = "| " & AlignL(A.S1, W1) & " | " & AlignL(A.S2, W2) & " |"
End Function
Function LinzDrvW$(Drv, WdtAy%())
Dim O$(), J%
For J = 0 To UB(Drv)
    PushI O, AlignL(CStr(Drv(J)), WdtAy(J))
Next
For J = UB(Drv) + 1 To UB(WdtAy)
    PushI O, Space(WdtAy(J))
Next
LinzDrvW = "| " & Jn(O, " | ") & " |"
End Function
Function Sy2zS1S2s(A As S1S2s) As String()
Dim J&
For J = 0 To A.N - 1
    PushI Sy2zS1S2s, CvS1S2(I).S2
Next
End Function

Private Function HasLines(A As S1S2s) As Boolean
Dim J&
HasLines = True
For J = 0 To UB(A)
    With A(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
HasLines = False
End Function


Private Sub Z_FmtS1S2s()
Dim Samp As New SampS1S2
Dim A As S1S2s, Nm1$, Nm2$
GoSub T0
'GoSub T1
'GoSub T2
Exit Sub
T0:
    Nm1 = "AA"
    Nm2 = "BB"
    PushS1S2
    A = AddS1S2(S1S2("A", "B"), S1S2("AA", "B"))
    GoTo Tst
T1:
    Nm1 = "AA"
    Nm2 = "BB"
    A = Samp.S1S2szLin
    GoTo Tst
T2:
    Nm1 = "AA"
    Nm2 = "BB"
    A = Samp.S1S2szLines
    GoTo Tst
Tst:
    Act = FmtS1S2s(A, Nm1, Nm2)
    BrwAy Act
    Return
End Sub

Private Sub Z()
Z_FmtS1S2s
End Sub
