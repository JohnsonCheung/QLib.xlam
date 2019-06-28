Attribute VB_Name = "QVb_Dta_S12_FmtS12"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_S12_Fmt."
Private Const Asm$ = "QVb"

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
      FmtS12s = FmtDrs(D, Fmt:=EiSSFmt, IxCol:=IxCol)
:         Exit Function
End If

Dim S1$():     S1 = S1Ay(A)
Dim S2$():     S2 = S2Ay(A)
Dim W1%:       W1 = WdtzLinesAy(AddElezStr(S1, N1))
Dim W2%:       W2 = WdtzLinesAy(AddElezStr(S2, N2))
Dim W2Ay%(): W2Ay = IntAy(W1, W2)
Dim SepL$:   SepL = LinzSep(W2Ay)
Dim Tit$:     Tit = AlignzDrWyAsLin(Array(N1, N2), W2Ay)
Dim M$():       M = XM(A, W2Ay, SepL)                    ' #Middle ! Middle part
Dim O$():       O = Sy(SepL, Tit, SepL, M)
               O = XAddIx(O, A.N, IxCol) '         ! Add Ix col in front

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
    Ly1 = AlignzAy(Ly1, W2Ay(0))
    Ly2 = AlignzAy(Ly2, W2Ay(1))
Dim O$()
    Dim J%, Dr(): For J = 0 To UB(Ly1)
        Dr = Array(Ly1(J), Ly2(J))
:            PushI O, AlignzDrWyAsLin(Dr, W2Ay)
    Next
XFmtzS12 = O
End Function

Private Function XM(A As S12s, W2Ay%(), SepL$) As String()
'Ret :  #Middle ! Middle part @@
Dim J&: For J = 0 To A.N - 1
    PushIAy XM, XFmtzS12(A.Ay(J), W2Ay)
    PushI XM, SepL
Next
'Insp "QVb_S1S2_Fmt.XM", "Inspect", "Oup(XM) A W2Ay SepL", XM, FmtS12s(A), W2Ay, SepL: Stop
End Function
Function AlignzDrWy(Dr, WdtAy%()) As String()
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
AlignzDrWy = O
End Function

Function AlignzDrWyAsLin$(Dr, WdtAy%())
'Ret : a lin by joing [ | ] and quoting [| * |] after aligng @Dr with @WdtAy. @@
AlignzDrWyAsLin = QteJnzAsTLin(AlignzDrWy(Dr, WdtAy))
End Function

Function S2Ay(A As S12s) As String()
Dim J&
For J = 0 To A.N - 1
    PushI S2Ay, A.Ay(J).S2
Next
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
Dim A As S12s, N1$, N2$
'GoSub T0
GoSub T1
GoSub T2
Exit Sub
T0:
    N1 = "AA"
    N2 = "BB"
    A = AddS12(S12("A", "B"), S12("AA", "B"))
    GoTo Tst
T1:
    N1 = "AA"
    N2 = "BB"
    A = SampS12szwLin
    GoTo Tst
T2:
    N1 = "AA"
    N2 = "BB"
    A = SampS12zwLines
    GoTo Tst
Tst:
    Act = FmtS12s(A, N1, N2)
    BrwAy Act
    Return
End Sub

