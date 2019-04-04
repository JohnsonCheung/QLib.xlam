Attribute VB_Name = "MVb_Thw_Msg"
Option Explicit

Function VblzLines$(Lines)
VblzLines = Replace(RmvCr(Lines), vbLf, "|")
End Function

Function LinzFunMsg$(Fun$, Msg$)
Dim F$: F = IIf(Fun = "", "", " | @" & Fun)
Dim A$: A = Msg & F
If Cfg_ShwInfLinTim Then
    LinzFunMsg = NowStr & " | " & A
Else
    LinzFunMsg = A
End If
End Function

Function LyzFunMsgNav(Fun$, Msg$, Nav()) As String()
LyzFunMsgNav = AyAdd(LyzFunMsg(Fun, Msg), AyIndent(LyzNav(Nav)))
End Function

Function LyzFunMsgNap(Fun$, Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
LyzFunMsgNap = LyzFunMsgNav(Fun, Msg, Nav)
End Function

Function LyzFunMsgObjPP(Fun$, Msg$, Obj, PP) As String()
LyzFunMsgObjPP = AyAdd(LyzFunMsg(Fun, Msg), LyzObjPP(Obj, PP))
End Function

Function LyzFunMsgNyAv(Fun$, Msg$, Ny$(), Av()) As String()
LyzFunMsgNyAv = AyAdd(LyzFunMsg(Fun, Msg), AyIndent(LyzNyAv(Ny, Av)))
End Function

Sub InfObjPP(Fun$, Msg$, Obj, PP)
D LyzFunMsgObjPP(Fun, Msg, Obj, PP)
End Sub

Function LyzNv(Nm$, V, Optional Sep$ = ": ") As String()
Dim Ly$(): Ly = LyzVal(V)
Dim J%, S$
If Si(Ly) = 0 Then
    PushI LyzNv, Nm & Sep
Else
    PushI LyzNv, Nm & Sep & Ly(0)
End If
S = Space(Len(Nm) + Len(Sep))
For J = 1 To UB(Ly)
    PushI LyzNv, S & Ly(J)
Next
End Function

Function LyzNvzStr$(Nm$, V)
LyzNvzStr = Nm & "=[" & StrCellzVal(V) & "]"
End Function
Function LyzMsgNap(Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
LyzMsgNap = LyzMsgNav(Msg, Nav)
End Function
Function LyzMsg(Msg$) As String()
LyzMsg = LyzFunMsg("", Msg)
End Function
Function LyzMsgNav(Msg$, Nav()) As String()
LyzMsgNav = AyAdd(LyzMsg(Msg), AyIndent(LyzNav(Nav)))
End Function

Function LyzNNAp(NN$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
LyzNNAp = LyzNNAv(NN, Av)
End Function

Function LyzNNAv(NN$, Av()) As String()
LyzNNAv = LyzNyAv(Ny(NN), Av)
End Function

Function LinzFunMsgNav$(Fun$, Msg$, Nav())
LinzFunMsgNav = LinzFunMsg(Fun, Msg) & " " & LinzNav(Nav)
End Function

Sub DmpNNAp(NN, ParamArray Ap())
Dim Av(): Av = Ap
D LyzNyAv(Ny(NN), Av)
End Sub

Function LyzNyAv(Ny$(), Av(), Optional Sep$ = ": ") As String()
Dim J%, O$(), N$()
ReSumSiabMax Ny, Av
N = FmtAySamWdt(Ny)
For J = 0 To UB(Ny)
    PushIAy LyzNyAv, LyzNv(N(J), Av(J), Sep)
Next
End Function

Function LinzLyzMsgNav$(Msg$, Nav())
LinzLyzMsgNav = SfxDotEns(Msg) & " | " & LinzNav(Nav)
End Function

Function LinzNav$(Nav())
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LinzNav = LinzNyAv(Ny, Av)
End Function

Function LinzNyAv$(Ny$(), Av())
Dim J%, U1%, U2%, N$, V$, O$()
U1 = UB(Ny)
U2 = UB(Av)
For J = 0 To Max(U1, U2)
    If J <= U1 Then N = QuoteSq(Ny(J)) Else N = "[?]"
    If J <= U2 Then V = StrCellzVal(Av(J)) Else V = "?"
    PushI O, N & " " & V
Next
LinzNyAv = JnVbarSpc(O)
End Function

Sub AsgNyAv(Nav(), ONy$(), OAv())
If Si(Nav) = 0 Then
    Erase ONy
    Erase OAv
    Exit Sub
End If
ONy = TermAyzTT(Nav(0))
OAv = AyeFstEle(Nav)
End Sub

Function LyzNav(Nav()) As String()
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LyzNav = LyzNyAv(Ny, Av)
End Function

Function SclzNyAv$(Ny$(), Av())
SclzNyAv = JnSemi(LyzNyAv(Ny, Av))
End Function

Function Box(S) As String()
Dim H$: H = Dup("*", Len(S) + 6)
PushI Box, H
PushI Box, "** " & S & " **"
PushI Box, H
End Function

Private Function LyzFunMsg(Fun$, Msg$) As String()
Dim O$(), MsgL1$, MsgRst$
AsgBrk1Dot Msg, MsgL1, MsgRst
PushI LyzFunMsg, SfxDotEns(MsgL1) & IIf(Fun = "", "", "  @" & Fun)
PushIAy LyzFunMsg, AyIndent(LyWrp(SplitCrLf(MsgRst)))
End Function


