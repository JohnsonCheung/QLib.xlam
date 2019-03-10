Attribute VB_Name = "MVb_S1S2_Fmt"
Option Explicit
Function FmtS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As String()
Dim mHasLines As Boolean: mHasLines = HasLines(A)
Dim mSepChr$:               mSepChr = IIf(mHasLines, "|", " ")
Dim mS1$():                     mS1 = Sy1zS1S2Ay(A)
Dim mS2$():                     mS2 = Sy2zS1S2Ay(A)
Dim mW1%:                       mW1 = WdtzLinesAy(AyAddItm(mS1, Nm1))
Dim mW2%:                       mW2 = WdtzLinesAy(AyAddItm(mS2, Nm2))
Dim mIncW%:                   mIncW = IIf(mHasLines, 2, 1)
Dim mSep$:                     mSep = SepLin(IntAy(mW1 + mIncW, mW2 + mIncW), mSepChr)
Dim mTit$:                     mTit = LinzS1S2(S1S2(Nm1, Nm2), mW1, mW2)
Dim mHdrLy$():               mHdrLy = Sy(mSep, mTit, mSep)
Dim mMidLy$():               mMidLy = LyzS1S2Ay(A, mW1, mW2, mHasLines, mSep)
FmtS1S2Ay = AyAdd(mHdrLy, mMidLy)
End Function
Function Sy1zS1S2Ay(A() As S1S2) As String()
Dim I
For Each I In Itr(A)
    PushI Sy1zS1S2Ay, CvS1S2(I).S1
Next
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
    ReszAyabMax Ly1, Ly2
    Ly1 = LyAlignLWdt(Ly1, W1)
    Ly2 = LyAlignLWdt(Ly2, W2)
Dim J%, O$()
For J = 0 To UB(Ly1)
    PushI O, "| " & Ly1(J) & " | " & Ly2(J) & " |"
Next
LyzS1S2 = O
End Function
Function LyAlignLWdt(A, W%) As String()
Dim I
For Each I In Itr(A)
    PushI LyAlignLWdt, AlignL(I, W)
Next
End Function
Private Function CvS1S2(A) As S1S2
Set CvS1S2 = A
End Function
Private Function LyzS1S2Ay(A() As S1S2, W1%, W2%, S1S2AyHasLines As Boolean, SepLin$) As String()
If S1S2AyHasLines Then
    Dim I
    For Each I In A
        PushIAy LyzS1S2Ay, LyzS1S2(CvS1S2(I), W1, W2)
        PushI LyzS1S2Ay, SepLin
    Next
Else
    For Each I In A
        PushI LyzS1S2Ay, LinzS1S2(CvS1S2(I), W1, W2)
    Next
    PushI LyzS1S2Ay, SepLin
End If
End Function
Private Function LinzS1S2$(A As S1S2, W1%, W2%)
LinzS1S2 = "| " & AlignL(A.S1, W1) & " | " & AlignL(A.S2, W2) & " |"
End Function
Function LinDrWdtAy$(Dr, WdtzAy%())
Dim O$(), J%
For J = 0 To UB(Dr)
    PushI O, AlignL(Dr(J), WdtzAy(J))
Next
For J = UB(Dr) + 1 To UB(WdtzAy)
    PushI O, Space(WdtzAy(J))
Next
LinDrWdtAy = "| " & Jn(O, " | ") & " |"
End Function
Function Sy2zS1S2Ay(A() As S1S2) As String()
Dim I
For Each I In Itr(A)
    PushI Sy2zS1S2Ay, CvS1S2(I).S2
Next
End Function

Private Function HasLines(A() As S1S2) As Boolean
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


Private Sub Z_FmtS1S2Ay()
Dim Samp As New SampS1S2
Dim A() As S1S2, Nm1$, Nm2$
GoSub T0
'GoSub T1
'GoSub T2
Exit Sub
T0:
    Nm1 = "AA"
    Nm2 = "BB"
    A = S1S2Ay(S1S2("A", "B"), S1S2("AA", "B"))
    GoTo Tst
T1:
    Nm1 = "AA"
    Nm2 = "BB"
    A = Samp.S1S2AyzLin
    GoTo Tst
T2:
    Nm1 = "AA"
    Nm2 = "BB"
    A = Samp.S1S2AyzLines
    GoTo Tst
Tst:
    Act = FmtS1S2Ay(A, Nm1, Nm2)
    BrwAy Act
    Return
End Sub

Private Sub Z()
Z_FmtS1S2Ay
End Sub
