Attribute VB_Name = "MVb_RunFil"
Option Explicit
Type WaitOpt
    TimOutSec As Integer
    ChkIntervalDeciSec As Integer
End Type

Function WaitOpt(TimOutSec%, ChkIntervalDeciSec%) As WaitOpt
With WaitOpt
.TimOutSec = TimOutSec
.ChkIntervalDeciSec = ChkIntervalDeciSec
End With
End Function

Property Get DftWaitOpt() As WaitOpt
DftWaitOpt = WaitOpt(30, 5)
End Property

Sub KillProcessId(ProcessId&)
StopXls
End Sub

Sub RunFcmd(Fcmd$, ParamArray PmAp())
Dim Av(): Av = PmAp
RunFcmdAv Fcmd, Av
End Sub

Function RunFcmdWait(Fcmd$, A As WaitOpt, ParamArray PmAp()) As Boolean
Dim Av(): Av = PmAp
Dim ProcessId&
ProcessId = RunFcmdAv(Fcmd, Av)
If Not WaitFfn(WaitFfnzFcmd(Fcmd)) Then
    KillProcessId ProcessId
    Exit Function
End If
RunFcmdWait = True
End Function

Function RunFcmdAv&(Fcmd, PmAv())
Dim Lin$
    Lin = JnSpc(AyQuoteDbl(AyItmAddAy(Fcmd, PmAv)))
RunFcmdAv = Shell(Lin, vbMaximizedFocus)
End Function

Private Sub ZZ_RunCmd()
RunFil "Cmd"
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
For J = 1 To TimOutSec \ ChkIntervalSec
    Wait
Next
End Function

Function WaitDeci(Optional DeciSec% = 10) As Boolean
WaitDeci = Xls.Wait(NxtDeciSec(DeciSec))
End Function

Function NxtDeciSec(DeciSec%) As Date
NxtDeciSec = DateAdd("S", DeciSec / 10, Now)
End Function
