Attribute VB_Name = "QVb_Str_Fmt"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Str_Fmt."
Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl$, Av())
Const CSub$ = CMod & "FmtQQAv"
Dim O$, I, Cnt
O = Replace(QQVbl, "|", vbCrLf)
Cnt = SubStrCnt(QQVbl, "?")
If Cnt <> Si(Av) Then
    Thw CSub, "[QQVbl-?-Cnt] <> Av-Si", "QQVbl-?-Cnt AvSz QQVbl Av", Cnt, Si(Av), QQVbl, Av
    Exit Function
End If
Dim P&
P = 1
For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Stop
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function

Private Sub ZZ_FmtQQAv()
Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
End Sub

Function LblTabFmtAySepSS(Lbl$, Sy$()) As String()
PushI LblTabFmtAySepSS, Lbl
PushIAy LblTabFmtAySepSS, TabSy(Sy)
End Function
