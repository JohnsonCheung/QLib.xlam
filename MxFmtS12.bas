Attribute VB_Name = "MxFmtS12"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFmtS12."

Private Function XDy(A As S12s) As Variant()
'Ret : a 2 col of dry with fst row is @N1..2 and snd row is ULin and rst from @A @@
Dim J&: For J& = 0 To A.N - 1
    With A.Ay(J)
    PushI XDy, Array(.S1, .S2)
    End With
Next
End Function

Function FmtS12s(A As S12s, Optional N1$ = "S1", Optional N2$ = "S2", Optional IxCol As EmIxCol) As String()
If A.N = 0 Then
    PushI FmtS12s, "(NoRec-S12s) (" & N1 & ") (" & N2 & ")"
    Exit Function
End If
Dim Dy(), D As Drs
If Not XHasLines(A) Then
          Dy = XDy(A)
           D = Drs(Sy(N1, N2), Dy)
     FmtS12s = FmtCellDrs(D, Fmt:=EiSSFmt, IxCol:=IxCol)
:              Exit Function
End If
Dim S1$():     S1 = S1Ay(A)
Dim S2$():     S2 = S2Ay(A)
Dim W1%:       W1 = WdtzLinesAy(AddEleS(S1, N1))
Dim W2%:       W2 = WdtzLinesAy(AddEleS(S2, N2))
Dim W2Ay%(): W2Ay = IntAy(W1, W2)
Dim SepL$:   SepL = LinzSep(W2Ay)
Dim Tit$:     Tit = AlignDrWyAsLin(Array(N1, N2), W2Ay)
Dim M$():       M = XMiddle(A, W2Ay, SepL)              '  #Middle ! Middle part
Dim O$():       O = Sy(SepL, Tit, SepL, M)
                O = XAddIx(O, A.N, IxCol)               '          ! Add Ix col in front

FmtS12s = O
End Function

Private Function XIxFront$(Fst2Chr$, IsIxAdd As Boolean, Sep$, Ix&, W%)
Dim O$
Select Case True
Case Fst2Chr = "|-":             O = Sep
Case Fst2Chr = "| " And IsIxAdd: O = "| " & Space(W + 1)
Case Fst2Chr = "| ":             O = "| " & Align(Ix, W) & " "
Case Else: Thw CSub, "Fst2Chr should [| ] or [|-]", "Fst2Chr", Fst2Chr
End Select
XIxFront = O
End Function

Private Function XAddIx(Fmt$(), N&, IxCol As EmIxCol) As String()
'@Fmt : ! a formatted S12s-Ly
'Ret  : ! Add Ix column in front of @Fmt @@
If IxCol = EiNoIx Then XAddIx = Fmt: Exit Function
Dim W%: W = Len(CStr(N))      ' AlignL width
Dim S$: S = "|" & Dup("-", W + 2) ' Sep lin
Dim IsIxAdd As Boolean            ' Is-Ix-Added.
Dim F$                            ' Front str to be added in front of each line
Dim F2$ ' Fst 2 chr of each lin of @Fmt
Dim Ix&: Ix = -1 ' The ix to be add
If IxCol = EiBeg1 Then Ix = Ix + 1
PushI XAddIx, S & Fmt(0)
PushI XAddIx, "| " & AlignR("#", W) + " " & Fmt(1)

Dim J&: For J = 2 To UB(Fmt)
        F2 = Fst2Chr(Fmt(J))
        If F2 = "|-" Then IsIxAdd = False: Ix = Ix + 1
    F = XIxFront(F2, IsIxAdd, S, Ix, W) 'What to add infront the a lin of @Fmt as an Ix col.
        If F2 = "| " And Not IsIxAdd Then IsIxAdd = True
        PushI XAddIx, F & Fmt(J)
Next
End Function

Private Function XFmtzS12(A As S12, W2Ay%()) As String()
'@A    : the :S12.S1-S2 may both have lines.  Wrap them as @W2Ay.
'@W2Ay : S1-Wdt and S2-Wdt
'Ret   : Ly aft fmt @A @@
Dim Ly1$(), Ly2$()
    Ly1 = SplitCrLf(A.S1)
    Ly2 = SplitCrLf(A.S2)
          ResiMax Ly1, Ly2
    Ly1 = AlignAy(Ly1, W2Ay(0))
    Ly2 = AlignAy(Ly2, W2Ay(1))
Dim O$()
    Dim J%, Dr(): For J = 0 To UB(Ly1)
        Dr = Array(Ly1(J), Ly2(J))
:            PushI O, AlignDrWyAsLin(Dr, W2Ay)
    Next
XFmtzS12 = O
End Function

Private Function XMiddle(A As S12s, W2Ay%(), SepL$) As String()
'Ret :  #Middle ! Middle part @@
Dim J&: For J = 0 To A.N - 1
    PushIAy XMiddle, XFmtzS12(A.Ay(J), W2Ay)
    PushI XMiddle, SepL
Next
'Insp "QVb_S1S2_Fmt.XMiddle", "Inspect", "Oup(XMiddle) A W2Ay SepL", XMiddle, FmtS12s(A), W2Ay, SepL: Stop
End Function
Function AlignDrWy(Dr, WdtAy%()) As String()
Dim O$()
Dim UDr&: UDr = UB(Dr)
Dim W, J%, S$: For Each W In WdtAy
    If J > UDr Then
        S = Space(W)
    Else
        S = AlignL(Dr(J), W)
    End If
    PushI O, S
    J = J + 1
Next
AlignDrWy = O
End Function


Private Function XHasLines(A As S12s) As Boolean
Dim J&
XHasLines = True
For J = 0 To A.N - 1
    With A.Ay(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
XHasLines = False
End Function

Private Sub Z_FmtS12s()
Dim A As S12s, N1$, N2$, Pseg$
'GoSub T0
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T3:
    N1 = "AA"
    N2 = "BB"
    Pseg = "Z_FmtS12s\Cas3"
    A = S12szRes("S12s.Txt", Pseg & "\Inp")
    Ept = ResStr("Ept", Pseg)
    GoTo Tst
T0:
    N1 = "AA"
    N2 = "BB"
    A = AddS12(S12("A", "B"), S12("AA", "B"))
    GoTo Tst
T1:
    N1 = "AA"
    N2 = "BB"
    A = SampS12s
    GoTo Tst
T2:
    N1 = "AA"
    N2 = "BB"
    A = SampS12s_wiLines
    Brw FmtS12s(A, N1, N2)
    Stop
    GoTo Tst
Tst:
    Act = FmtS12s(A, N1, N2)
    C
    Return
End Sub
