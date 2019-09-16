Attribute VB_Name = "MxMsg"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxMsg."
Function AddNNAv(Nav(), Nn$, Av()) As Variant()
Dim O(): O = Nav
If Si(O) = 0 Then
    PushI O, Nn
Else
    O(0) = O(0) & " " & Nn
End If
PushAy O, Av
AddNNAv = O
End Function

Function AddNmV(Nav(), NM$, V) As Variant()
AddNmV = AddNNAv(Nav, NM, Av(V))
End Function


Sub InfLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) >= 0 Then Nav = Nap
D LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub InfNav(Fun$, Msg$, Nav())
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Inf(Fun$, Msg$, ParamArray Nap())
If Not Cfg.Inf.ShwInf Then Exit Sub
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub WarnLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Debug.Print LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Warn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub DmpNNAp(Nn$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
D LyzNyAv(Ny(Nn), Av)
End Sub

Function LyzNyAv(Ny$(), Av(), Optional Sep$ = ": ") As String()
Dim J%, O$(), N$()
ResiMax Ny, Av
N = AlignAy(Ny)
For J = 0 To UB(Ny)
    PushIAy LyzNyAv, LyzNv(N(J), Av(J), Sep)
Next
End Function

Function LinzLyzMsgNav$(Msg$, Nav())
LinzLyzMsgNav = EnsSfxDot(Msg) & " | " & LinzNav(Nav)
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
    If J <= U1 Then N = QteSq(Ny(J)) Else N = "[?]"
    If J <= U2 Then V = Cell(Av(J)) Else V = "?"
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
Dim TT$: TT = Nav(0)
ONy = TermAy(TT)
OAv = AeFstEle(Nav)
End Sub
Private Sub Z_LyzNav()
Dim Nav(): Nav = Array("aa bb", 1, 2)
D LyzNav(Nav)
End Sub

Function LyzNav(Nav()) As String()
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LyzNav = LyzNyAv(Ny, Av)
End Function

Function SclzNyAv$(Ny$(), Av())
SclzNyAv = JnSemi(LyzNyAv(Ny, Av))
End Function

Function LyzFunMsg(Fun$, Msg$) As String()
Dim O$(), MsgL1$, MsgRst$
AsgBrk1Dot Msg, MsgL1, MsgRst
PushI LyzFunMsg, EnsSfxDot(MsgL1) & IIf(Fun = "", "", "  @" & Fun)
PushIAy LyzFunMsg, IndentSy(WrpLy(SplitCrLf(MsgRst)))
End Function

Function LyzFunMsgNav(Fun$, Msg$, Nav()) As String()
Dim A$(): A = LyzFunMsg(Fun, Msg)
Dim B$(): B = IndentSy(LyzNav(Nav))
LyzFunMsgNav = AddAy(A, B)
End Function

Function LyzFunMsgNap(Fun$, Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
If Fun = "" And Msg = "" And Si(Nav) = 0 Then Exit Function
LyzFunMsgNap = LyzFunMsgNav(Fun, Msg, Nav)
End Function

Function LyzFunMsgObjPP(Fun$, Msg$, Obj As Object, PP$) As String()
LyzFunMsgObjPP = AddAy(LyzFunMsg(Fun, Msg), LyzObjPP(Obj, PP))
End Function

Function LyzFunMsgNyAv(Fun$, Msg$, Ny$(), Av()) As String()
LyzFunMsgNyAv = AddAy(LyzFunMsg(Fun, Msg), IndentSy(LyzNyAv(Ny, Av)))
End Function

Function LyzFunMsgNNAv(Fun$, Msg$, Nn$, Av()) As String()
LyzFunMsgNNAv = LyzFunMsgNyAv(Fun, Msg, SyzSS(Nn), Av)
End Function

Function LyzNv(NM$, V, Optional Sep$ = ": ") As String()
Dim Ly$(): Ly = FmtV(V)
Dim J%, S$
If Si(Ly) = 0 Then
    PushI LyzNv, NM & Sep
Else
    PushI LyzNv, NM & Sep & Ly(0)
End If
S = Space(Len(NM) + Len(Sep))
For J = 1 To UB(Ly)
    PushI LyzNv, S & Ly(J)
Next
End Function

Function LinzNv$(NM$, V)
LinzNv = NM & "=[" & Cell(V) & "]"
End Function

Function LyzMsgNap(Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
LyzMsgNap = LyzMsgNav(Msg, Nav)
End Function

Function LyzNmDrs(NM$, A As Drs, Optional MaxColWdt% = 100) As String()
LyzNmDrs = LyzNmLy(NM, FmtCellDrs(A, MaxColWdt), EiNoIx)
End Function

Function LyzNmLy(NM$, Ly$(), Optional B As EmIxCol = EiBeg1) As String()
Dim L$(), J&, S$
If Si(Ly) = 0 Then
    PushI LyzNmLy, NM & "(No Lin)"
    Exit Function
End If
L = AddIxPfx(Ly, B)
'Brw L:Stop
S = Space(Len(NM))
PushI LyzNmLy, NM & L(0)
For J = 1 To UB(L)
    PushI LyzNmLy, S & L(J)
Next
End Function

Function LyzMsg(Msg$) As String()
LyzMsg = LyzFunMsg("", Msg)
End Function
Function LyzMsgNav(Msg$, Nav()) As String()
LyzMsgNav = AddAy(LyzMsg(Msg), IndentSy(LyzNav(Nav)))
End Function

Function LyzNNAp(Nn$, ParamArray Ap()) As String()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
LyzNNAp = LyzNNAv(Nn, Av)
End Function



Function LyzNNAv(Nn$, Av()) As String()
LyzNNAv = LyzNyAv(Ny(Nn), Av)
End Function

Function LinzFunMsgNav$(Fun$, Msg$, Nav())
LinzFunMsgNav = LinzFunMsg(Fun, Msg) & " " & LinzNav(Nav)
End Function

Function LinzFunMsg$(Fun$, Msg$)
Dim F$: F = IIf(Fun = "", "", " | @" & Fun)
Dim A$: A = Msg & F
If Cfg.Inf.ShwTim Then
    LinzFunMsg = NowStr & " | " & A
Else
    LinzFunMsg = A
End If
End Function