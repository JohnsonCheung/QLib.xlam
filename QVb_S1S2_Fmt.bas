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
Function DryzInsIx(Dry()) As Variant()
' Ret Dry with each row has ix run from 0..{N-1} in front
Dim Ix&, Dr: For Each Dr In Itr(Dry)
    InsEle Dr, Ix
Next
End Function

Private Function XDry(A As S1S2s, N1$, N2$) As Variant()
'Ret : a 2 col of dry with fst row is @N1..2 and snd row is ULin and rst from @A @@
PushI XDry, Array(N1, N2)
PushI XDry, Array(ULin(N1), ULin(N2))
Dim J&: For J& = 0 To A.N - 1
    With A.Ay(J)
    PushI XDry, Array(.S1, .S2)
    End With
Next
End Function

Function FmtS1S2s(A As S1S2s, Optional N1$ = "S1", Optional N2$ = "S2", Optional SkipIx As Boolean) As String()
If A.N = 0 Then
    PushI FmtS1S2s, "(NoRec-S1S2s) (" & N1 & ") (" & N2 & ")"
    Exit Function
End If
If Not XHasLines(A) Then
    Dim Dry(): Dry = XDry(A, N1, N2)
                     If Not SkipIx Then Dry = DryzInsIx(Dry)
    FmtS1S2s = AlignzDryAsLy(Dry)
    Exit Function
End If

Dim S1$():    S1 = S1Ay(A)
Dim S2$():    S2 = S2Ay(A)
Dim W1%:      W1 = WdtzLinesAy(AddElezStr(S1, N1))
Dim W2%:      W2 = WdtzLinesAy(AddElezStr(S2, N2))
Dim W2Ay%(): W2Ay = IntAy(W1, W2)
Dim SepL$:   SepL = LinzSep(W2Ay)
Dim Tit$:     Tit = AlignzDrWyAsLin(Array(N1, N2), W2Ay)
Dim M$():       M = XM(A, W2Ay, SepL) ' #Middle ! Middle part
Dim O$():       O = Sy(SepL, Tit, SepL, M)
                    If Not SkipIx Then O = XAddIx(O, A.N) ' ! Add Ix col in front
         FmtS1S2s = O
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

Private Function XAddIx(Fmt$(), N&) As String()
'Fm Fmt : ! a formatted
'Ret : Add Ix column in front of @Fmt
Dim W%: W = Len(CStr(N))      ' AlignL width
Dim S$: S = "|" & Dup("-", W + 2) ' Sep lin
Dim IsIxAdd As Boolean            ' Is-Ix-Added.
Dim F$                            ' Front str to be added in front of each line
Dim F2$ ' Fst 2 chr of each lin of @Fmt
Dim Ix&: Ix = -1 ' The ix to be add
PushI XAddIx, S & Fmt(0)
PushI XAddIx, "| " & Align("#", W) + " " & Fmt(1)

Dim J&: For J = 2 To UB(Fmt)
        F2 = Fst2Chr(Fmt(J))
        If F2 = "|-" Then IsIxAdd = False: Ix = Ix + 1
    F = XIxFront(F2, IsIxAdd, S, Ix, W) 'What to add infront the a lin of @Fmt as an Ix col.
        If F2 = "| " And Not IsIxAdd Then IsIxAdd = True
        PushI XAddIx, F & Fmt(J)
Next
End Function

Private Function XLyzS1S2(A As S1S2, W2Ay%()) As String()
Dim Ly1$(), Ly2$()
    Ly1 = SplitCrLf(A.S1)
    Ly2 = SplitCrLf(A.S2)
          ResiMax Ly1, Ly2
    Ly1 = AlignzAy(Ly1, W2Ay(0))
    Ly2 = AlignzAy(Ly2, W2Ay(1))
Dim J%, O$(): For J = 0 To UB(Ly1)
    Dim Dr(): Dr = Array(Ly1(J), Ly2(J))
                   PushI O, AlignzDrWyAsLin(Dr, W2Ay)
Next
XLyzS1S2 = O
End Function

Private Function XM(A As S1S2s, W2Ay%(), SepL$) As String()
Dim J&: For J = 0 To A.N - 1
    PushIAy XM, XLyzS1S2(A.Ay(J), W2Ay)
    PushI XM, SepL
Next
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

Function S2Ay(A As S1S2s) As String()
Dim J&
For J = 0 To A.N - 1
    PushI S2Ay, A.Ay(J).S2
Next
End Function

Private Function XHasLines(A As S1S2s) As Boolean
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

Sub Z_FmtS1S2s()
Dim A As S1S2s, N1$, N2$
'GoSub T0
GoSub T1
GoSub T2
Exit Sub
T0:
    N1 = "AA"
    N2 = "BB"
    A = AddS1S2(S1S2("A", "B"), S1S2("AA", "B"))
    GoTo Tst
T1:
    N1 = "AA"
    N2 = "BB"
    A = SampS1S2szwLin
    GoTo Tst
T2:
    N1 = "AA"
    N2 = "BB"
    A = SampS1S2zwLines
    GoTo Tst
Tst:
    Act = FmtS1S2s(A, N1, N2, SkipIx:=False)
    BrwAy Act
    Return
End Sub

Sub Z()
Z_FmtS1S2s
End Sub
