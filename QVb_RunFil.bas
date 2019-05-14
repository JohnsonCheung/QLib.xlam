Attribute VB_Name = "QVb_RunFil"
Option Explicit
Private Const CMod$ = "MVb_RunFil."
Private Const Asm$ = "QVb"
Enum EmWaitRslt
    EiTimUp
    EiCnl
End Enum
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
Dim Lin
    Lin = QuoteDbl(Fcmd) & PpdSpcIf(PmStr)
RunFcmd = Shell(Lin, Sty)
End Function

Function WaitFcmdw(Fcmdw$, W As WaitOpt, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fcmdw has generated the Fwaitg
Dim ProcessId&: ProcessId = Shell(Fcmdw, Sty)
Dim Fw$: Fw = Fwaitg(Fcmdw)
If WaitFwaitg(Fw, W.ChkIntervalDeciSec, W.TimOutSec) Then
    Kill Fw
    WaitFcmdw = True
Else
    KillProcessId ProcessId
End If
If Not W.KeepFcmd Then Kill Fcmdw
End Function

Private Function WaitFwaitg(Fwaitg$, Optional ChkIntervalDeciSec% = 10, Optional TimOutSec% = 60, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus) As Boolean _
'Return True, if Fwaitg is found.
Dim J%
For J = 1 To TimOutSec \ ChkIntervalDeciSec
    If HasFfn(Fwaitg) Then
        Kill Fwaitg
        Exit Function
    End If
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
Dim EchoLin: EchoLin = FmtQQ("Echo > ""?""", Fwaitg(T))
Dim S$: S = CmdLines & vbCrLf & EchoLin
Fcmdw = WrtStr(S, T)
End Function

Private Sub ZZ_RunFcmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As EmWaitRslt
Dim Till As Date: Till = AftSec(Sec)
Wait = IIf(Xls.Wait(Till), EiTimUp, EiCnl)
End Function

Function AftSec(Sec%) As Date 'Return the Date after Sec from Now
AftSec = DateAdd("S", Sec, Now)
End Function
