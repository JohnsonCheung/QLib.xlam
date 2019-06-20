Attribute VB_Name = "QVb_Str_Expand"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Expand."
Private Const Asm$ = "QVb"
Function Expand$(QVblTp$, Seed$())
Dim O$(), Tp$, ISeed
Tp = RplVBar(QVblTp)
For Each ISeed In Itr(Seed)
    PushI O, SzQBy(Tp, ISeed)
Next
Expand = Join(O, "")
End Function

Function Expandss$(QVblTp$, Seedss$)
Expandss = Expand(QVblTp, SyzSS(Seedss))
End Function
Private Sub Z_Expandss()
Dim QVblTp$, Seed$()
Z:
    Erase XX
    X "Sub Push?(O() As ?, M As ?)"
    X "Dim N&"
    X "N = ?Si(O)"
    X "ReDim Preserve O(N)"
    X "O(N) = M"
    X "End Sub"
    X ""
    X "Function ?Si&(A() As ?)"
    X "On Error Resume Next"
    X "?Si = Ubound(A) + 1"
    X "End Function"
    X ""
    X ""
    QVblTp = JnVBar(XX)
    Erase XX
    Brw Expandss(QVblTp, "S1S2 XX")
T0:
    QVblTp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
    Seed = SyzSS("Xws Xwb Xfx Xrg")
    Erase XX
    X ""
    X ""
    Ept = JnCrLf(XX)
    Erase XX
    GoTo Tst
Tst:
    Act = Expand(QVblTp, Seed)
    C
    Return
End Sub

Private Sub Z()
Z_Expandss
End Sub

