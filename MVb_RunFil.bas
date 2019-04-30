Attribute VB_Name = "MVb_RunFil"
Option Explicit
Type WaitOpt
    TimOutSec As Integer
    ChkIntervalDeciSec As Integer
    KeepFcmd As Boolean
End Type

Function WaitOpt(TimOutSec%, ChkIntervalDeciSec%, KeepFcmd As Boolean) As WaitOpt
With WaitOpt
.TimOutSec = TimOutSec
.ChkIntervalDeciSec = ChkIntervalDeciSec
.KeepFcmd = KeepFcmd
End With
End Function

Property Get DftWaitOpt() As WaitOpt
DftWaitOpt = WaitOpt(30, 5, False)
End Property

Sub KillProcessId(ProcessId&)
End Sub

Function RunFcmd&(Fcmd$, Optional PmStr$, Optional Sty As VbAppWinStyle = vbMaximizedFocus)
Dim Lin$
    Lin = QuoteDbl(Fcmd) & PpdSpcIf(PmStr)
RunFcmd = Shell(Lin, Sty)
End Function

Private Function WaitFcmd(Fcmd$, W As WaitOpt, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbHide) As Boolean
Dim Fw$: Fw = Fcmdw(Fcmd)
Dim ProcessId&
ProcessId = Shell(Fw, Sty)
If Not WaitFcmdw(Fw, W.ChkIntervalDeciSec, W.TimOutSec) Then
    KillProcessId ProcessId
    Exit Function
End If
If Not W.KeepFcmd Then Kill Fcmd
WaitFcmd = True
End Function

Function WaitFcmdw(Fcmdw$, Optional ChkIntervalDeciSec% = 10, Optional TimOutSec% = 60) As Boolean 'Return True, if Fwaitg is found.
Dim J%
Dim Fw$: Fw = Fwaitg(Fcmdw)
For J = 1 To TimOutSec \ ChkIntervalDeciSec
    If HasFfn(Fw) Then WaitFcmdw = True: Exit Function
    If Not Wait(ChkIntervalDeciSec%) Then Exit Function
Next
End Function

Private Sub Z_Fcmdw()
Debug.Print LineszFt(Fcmdw("Dir"))
End Sub
Function Fwaitg$(Fcmd$)
Fwaitg = Fcmd & ".wait.txt"
End Function

Function Fcmdw$(CmdLines$)
Dim T$: T = TmpCmd
Dim EchoLin$: EchoLin = FmtQQ("Echo > ""?""", Fwaitg(T))
Dim S$: S = CmdLines & vbCrLf & EchoLin
Fcmdw = WrtStr(S, T)
End Function

Private Sub ZZ_RunFcmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As Boolean
Wait = WaitDeci(Sec / 10)
End Function

Function WaitDeci(Optional DeciSec% = 10) As Boolean
WaitDeci = Xls.Wait(NxtDeciSec(DeciSec))
End Function

Function NxtDeciSec(DeciSec%) As Date
NxtDeciSec = DateAdd("S", DeciSec / 10, Now)
End Function
