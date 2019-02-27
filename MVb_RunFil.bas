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

Sub RunFcmd(Fcmd$, ParamArray PmAp())
Dim Av(): Av = PmAp
RunFcmdAv Fcmd, Av
End Sub
Private Function RunFcmdWaitOpt(Fcmd$, A As WaitOpt, ParamArray PmAp()) As Boolean
Dim Av(): Av = PmAp
RunFcmdWaitOpt = RunFcmdWaitOptAv(Fcmd, A, Av)
End Function

Private Function RunFcmdWaitOptAv(Fcmd$, A As WaitOpt, ParamArray PmAp()) As Boolean
Dim ProcessId&
ProcessId = RunFcmdAv(Fcmd, Av)
If Not WaitFfn(WaitFfnzFcmd(Fcmd)) Then
    KillProcessId ProcessId
    Exit Function
End If
Kill WaitFfnzFcmd(Fcmd)
If Not A.KeepFcmd Then Kill Fcmd
RunFcmdWaitOptAv = True
End Function

Function RunFcmdWait(Fcmd$, ParamArray PmAp()) As Boolean
Dim Av(): Av = PmAp
RunFcmdWait = RunFcmdWaitOpt(Fcmd, DftWaitOpt, Av)
End Function

Function RunFcmdAv&(Fcmd, PmAv())
Dim Lin$
    Lin = JnSpc(AyQuoteDbl(AyItmAddAy(Fcmd, PmAv)))
RunFcmdAv = Shell(Lin, vbMaximizedFocus)
End Function

Private Sub ZZ_RunCmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function WaitFfnzFcmd$(Ffn)
WaitFfnzFcmd = Ffn & ".wait.txt"
End Function

Function Wait(Optional Sec% = 1) As Boolean
Wait = WaitDeci(Sec * 10)
End Function

Function WaitFfn(Ffn, Optional ChkIntervalDeciSec% = 10, Optional TimOutSec% = 60) As Boolean
Dim J%
For J = 1 To TimOutSec \ ChkIntervalDeciSec
    If HasFfn(Ffn) Then WaitFfn = True: Exit Function
    If Not Wait Then Exit Function
Next
End Function

Function WaitDeci(Optional DeciSec% = 10) As Boolean
WaitDeci = Xls.Wait(NxtDeciSec(DeciSec))
End Function

Function NxtDeciSec(DeciSec%) As Date
NxtDeciSec = DateAdd("S", DeciSec / 10, Now)
End Function
