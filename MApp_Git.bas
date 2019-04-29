Attribute VB_Name = "MApp_Git"
Option Explicit
Const CMod$ = "MApp_Git."
Public Const DocOfFwcmd$ = "It is a TmpFfn and the content is given-CmdLines plus EchoLin which is creating a WaitgFfn."
Public Const DocOfWaitgFfn$ = "WaitgFfn is a temp file without any content.  It is created at end of the Fwcmd."

Function CmitgFcmd$(Optional Msg$ = "commit", Optional ReInit As Boolean)
CmitgFcmd = Fwcmd(CdLineszGitCmit(Srcp(CurPj), Msg, ReInit))
End Function

Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Stop
WaitFfn CmitgFcmd(Msg, ReInit)
End Sub

Sub GitPush()
RunFcmdWait Fwcmd(GitPushCdLines(Srcp(CurPj)))
End Sub

Private Function CdLineszGitCmit$(CmitgPth$, Msg$, ReInit As Boolean)
Erase XX
Dim Pj$: Pj = PjNmzCmitgPth(CmitgPth)
X "Cd """ & CmitgPth & """"
If ReInit Then X "Rd .git /s/q"
X "git init"        'If already init, it will do nothing
X "git add -A"
X FmtQQ("git commit -m ""?""", Msg)
X "Pause"
CdLineszGitCmit = JnCrLf(XX)
Erase XX
End Function

Private Sub Z_Fwcmd()
Debug.Print LineszFt(Fwcmd("Dir"))
End Sub
Function WaitgFfn$(Fcmd$)
WaitgFfn = Fcmd & ".wait.txt"
End Function
Private Function Fwcmd$(CmdLines$)
Dim T$: T = TmpCmd
Dim EchoLin$: EchoLin = FmtQQ("Echo > ""?""", WaitgFfn(T))
Dim S$: S = CmdLines & vbCrLf & EchoLin
Fwcmd = WrtStr(S, T)
End Function

Function HasInternet() As Boolean
Stop
End Function

Private Function GitPushCdLines$(CmitgPth)
Dim O$(), Cd$, GitPush, T
Push O, FmtQQ("Cd ""?""", CmitgPth)
Push O, FmtQQ("git push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", PjNm(CmitgPth))
Push O, "Pause"
GitPushCdLines = JnCrLf(O)
End Function

Sub BrwGitCmitCdLines()
BrwFt CdLineszGitCmit("PthA", "Msg", ReInit:=True)
End Sub

Sub BrwGitPushCdLines()
BrwFt GitPushCdLines("A")
End Sub
Private Function PjNmzCmitgPth$(CmitgPth$)
Const CSub$ = CMod & "PjNm"
If Fdr(ParPth(CmitgPth)) <> ".Src" Then Thw CSub, "Not source path", "CmitgPth", CmitgPth
PjNmzCmitgPth = Fdr(CmitgPth)
End Function

Private Sub XX1()
'…or create a new repository on the command line
'echo "# QLib.xlam" >> README.md
'git Init
'git add README.md
'git commit -m "first commit"
'git remote add origin https://github.com/JohnsonCheung/QLib.xlam.git
'git push -u origin master
End Sub

Private Sub Z()
MApp_Commit:
End Sub

Sub PowerRun(Ps1, ParamArray PmAp())
Dim Av(): Av = PmAp
RunFcmdAv "PowerShell", Av
End Sub

