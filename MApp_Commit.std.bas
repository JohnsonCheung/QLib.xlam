Attribute VB_Name = "MApp_Commit"
Option Explicit
Sub Cmit(Optional Msg$ = "Commit")
CmitzPth SrcPthPj, Msg
End Sub

Private Sub CmitzPth(CmitgPth, Msg$)
EnsGitPth CmitgPth
RunFcmdWait CmitFcmd(CmitgPth, Msg), DftWaitOpt
End Sub

Private Sub EnsGitPth(CmitgPth)
If HasGit(CmitgPth) Then Exit Sub
RunFcmdWait InitFcmd(CmitgPth), DftWaitOpt
End Sub

Private Function HasGit(Pth) As Boolean
HasGit = HasPth(AddFdr(Pth, ".git"))
End Function

Private Property Get GitPushFcmd$()
GitPushFcmd = TmpPth & "Pushing.Cmd"
End Property

Sub GitPushApp()
EnsGitGitPushFcmdCxt
RunFcmdWait GitPushFcmdCxt, DftWaitOpt
End Sub

Private Function InitFcmd$(CmitgPth)
Dim T$: T = TmpCmd
InitFcmd = WrtStr(InitCd(CmitgPth, T), T)
End Function

Private Function InitCd$(CmitgPth, Fcmd$)
Erase XX
X "Cd """ & CmitgPth & """"
X "Git Init"
X "Echo Time >""" & WaitFfnzFcmd(Fcmd) & """"
InitCd = JnCrLf(XX)
Erase XX
End Function

Private Function CmitFcmd$(CmitgPth, Msg$)
Dim T$: T = TmpCmd
CmitFcmd = WrtStr(CmitCd(CmitgPth, Msg, T), T)
End Function

Private Sub EnsGitGitPushFcmdCxt()
WrtStr GitPushFcmdCxt, GitPushFcmd
End Sub

Private Function CmitCd$(CmitgPth, Msg$, Fcmd$)
Erase XX
X FmtQQ("Cd ""?""", CmitgPth)
X "git add -A"
X FmtQQ("git commit --message=""?""", Msg)
X "Echo Time >""" & WaitFfn(Fcmd) & """"
CmitCd = JnCrLf(XX)
Erase XX
End Function

Function HasInternet() As Boolean
Stop
End Function

Private Property Get GitPushFcmdCxt$()
Dim O$(), Cd$, GitPush, T
'Cd = FmtQQ("Cd ""?""", SrcPth)
Push O, Cd
Push O, "git push -u https://johnsoncheung@github.com/johnsoncheung/StockShipRate.git master"
Push O, "Pause"
GitPushFcmdCxt = JnCrLf(O)
End Property

Sub Exp()
ExpzPj CurPj
End Sub

Sub BrwInitFcmd()
BrwFt InitFcmd("PthA")
End Sub

Sub BrwCmitFcmd()
BrwFt CmitFcmd("PthA", "Commit")
End Sub

Sub BrwGitPushFcmd()
EnsGitGitPushFcmdCxt
BrwFt GitPushFcmd
End Sub

Private Sub XXX()
Dim O$()
Push O, "echo ""# QLib"" >> README.md"
Push O, "git Init"
Push O, "git add README.md"
Push O, "git commit -m ""first commit"""
Push O, "git remote add origin https://github.com/JohnsonCheung/QLib.git"
Push O, "git push -u origin master"

End Sub

Private Sub XX1()
'git remote add origin https://github.com/JohnsonCheung/QLib.git
'git push -u origin master
End Sub

Private Sub Z()
MApp_Commit:
End Sub

Sub PowerRun(Ps1, ParamArray PmAp())
Dim Av(): Av = PmAp
RunFcmdAv "PowerShell", Av
End Sub

