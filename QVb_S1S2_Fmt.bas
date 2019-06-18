Attribute VB_Name = "QVb_S1S2_Fmt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_S1S2_Fmt."
Private Const Asm$ = "QVb"
Function AddS1S2(A As S1S2, B As S1S2) As S1S2s
Dim O As S1S2s
PushS1S2 O, A
PushS1S2 O, B
AddS1S2 = O
End Function
Function S1S2szS1S2(S1, S2) As S1S2s
Dim O As S1S2s
PushS1S2 O, S1S2(S1, S2)
S1S2szS1S2 = O
End Function
Function AddS1Pfx(A As S1S2s, S1Pfx$) As S1S2s
Dim J&: For J = 0 To A.N - 1
    Dim M As S1S2: M = A.Ay(J)
    M.S1 = S1Pfx & M.S1
    PushS1S2 AddS1Pfx, M
Next
End Function
Sub PushS1S2s(O As S1S2s, A As S1S2s)
Dim J&
For J = 0 To A.N - 1
    PushS1S2 O, A.Ay(J)
Next
End Sub
Function FmtS1S2s(A As S1S2s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As String()
If A.N = 0 Then
    PushI FmtS1S2s, "(NoRec-S1S2s) (" & Nm1 & ") (" & Nm2 & ")"
    Exit Function
End If
Dim mHasLines As Boolean, mSepChr$, mS1$(), mS2$(), mW1%, mW2%, mSepLin$, mHdrLy$(), mMidLy$()
mHasLines = HasLines(A)
      mS1 = Sy1zS1S2s(A)
      mS2 = Sy2zS1S2s(A)
      mW1 = WdtzLinesAy(AddElezStr(mS1, Nm1))
      mW2 = WdtzLinesAy(AddElezStr(mS2, Nm2))
  mSepLin = SepLin(IntAy(mW1, mW2))
   mHdrLy = HdrLy(mSepLin, Nm1, Nm2, mW1, mW2)
   mMidLy = LyzS1S2s(A, mW1, mW2, mHasLines, mSepLin)
FmtS1S2s = SyzAdd(mHdrLy, mMidLy)
End Function
Private Function HdrLy(mSep$, Nm1$, Nm2$, W1%, W2%) As String()
If Nm1 = "" And Nm2 = "" Then Exit Function
Dim mTit$:  mTit = LinzS1S2(S1S2(Nm1, Nm2), W1, W2)
HdrLy = Sy(mSep, mTit, mSep)
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
    ResiMax Ly1, Ly2
    Ly1 = SyzAlign(Ly1, W1)
    Ly2 = SyzAlign(Ly2, W2)
Dim J%, O$()
For J = 0 To UB(Ly1)
    PushI O, "| " & Ly1(J) & " | " & Ly2(J) & " |"
Next
LyzS1S2 = O
End Function

Private Function LyzS1S2s(A As S1S2s, W1%, W2%, S1S2sHasLines As Boolean, SepLin) As String()
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
    PushI O, AlignL(Drv(J), WdtAy(J))
Next
For J = UB(Drv) + 1 To UB(WdtAy)
    PushI O, Space(WdtAy(J))
Next
LinzDrvW = "| " & Jn(O, " | ") & " |"
End Function
Function Sy2zS1S2s(A As S1S2s) As String()
Dim J&
For J = 0 To A.N - 1
    PushI Sy2zS1S2s, A.Ay(J).S2
Next
End Function

Private Function HasLines(A As S1S2s) As Boolean
Dim J&
HasLines = True
For J = 0 To A.N - 1
    With A.Ay(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
HasLines = False
End Function


Private Sub Z_FmtS1S2s()
Dim A As S1S2s, Nm1$, Nm2$
GoSub T0
GoSub T1
GoSub T2
Exit Sub
T0:
    Nm1 = "AA"
    Nm2 = "BB"
    A = AddS1S2(S1S2("A", "B"), S1S2("AA", "B"))
    GoTo Tst
T1:
    Nm1 = "AA"
    Nm2 = "BB"
    A = SampS1S2szwLin
    GoTo Tst
T2:
    Nm1 = "AA"
    Nm2 = "BB"
    A = SampS1S2zwLines
    GoTo Tst
Tst:
    Act = FmtS1S2s(A, Nm1, Nm2)
    BrwAy Act
    Return
End Sub

Private Sub ZZ()
Z_FmtS1S2s
End Sub
