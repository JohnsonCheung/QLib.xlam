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

Sub RunFcmd(Fcmd$, Optional PmStr$)
Dim Lin$
PpdSpc
    Lin = JnSpc(SyQuoteDbl(StrAddSy(Fcmd, Pm)))
RunFcmdzPm = Shell(Fcmd & Lin, vbMaximizedFocus)
End Sub

Private Function WaitFwcmd(Fwcmd$, W As WaitOpt, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbHide) As Boolean
Dim ProcessId&
ProcessId = Shell(Fwcmd, Sty)
If Not WaitFfn(WaitgFfn(Fcmd)) Then
    KillProcessId ProcessId
    Exit Function
End If
Kill WaitgFfn(Fcmd)
If Not W.KeepFcmd Then Kill Fcmd
WaitFcmdPmAp = True
End Function


Private Sub ZZ_RunCmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As Boolean
Wait = WaitDeci(Sec * 10)
End Function

Function WaitFfn(Ffn$, Optional ChkIntervalDeciSec% = 10, Optional TimOutSec% = 60) As Boolean
Dim J%
For J = 1 To TimOutSec \ ChkIntervalDeciSec
    If HasFfn(Ffn$) Then WaitFfn = True: Exit Function
    If Not Wait Then Exit Function
Next
End Function

Function WaitDeci(Optional DeciSec% = 10) As Boolean
WaitDeci = Xls.Wait(NxtDeciSec(DeciSec))
End Function

Function NxtDeciSec(DeciSec%) As Date
NxtDeciSec = DateAdd("S", DeciSec / 10, Now)
End Function
