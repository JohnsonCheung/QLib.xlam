Attribute VB_Name = "MApp_Git"
Option Explicit

Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
RunFcmdWait FcmdWaitzCdLines(GitCmitCdLines(SrcPth(CurPj), Msg, ReInit))
End Sub

Sub GitPush()
RunFcmdWait FcmdWaitzCdLines(GitPushCdLines(SrcPth(CurPj)))
End Sub

Private Function GitCmitCdLines$(CmitgPth, Msg$, ReInit As Boolean)
Erase XX
Dim Pj$: Pj = PjNm(CmitgPth)
X "Cd """ & CmitgPth & """"
If ReInit Then X "Rd .git /s/q"
X "git init"
X "git add -A"
X FmtQQ("git commit -m ""?""", Msg)
X "Pause"
GitCmitCdLines = JnCrLf(XX)
Erase XX
End Function
Private Sub Z_FcmdWaitzCdLines()
Debug.Print FtLines(FcmdWaitzCdLines("Dir"))
End Sub
Private Function FcmdWaitzCdLines$(CdLines)
Dim T$: T = TmpCmd
Dim EchoLin$: EchoLin = FmtQQ("Echo > ""?""", WaitFfnzFcmd(T))
Dim S$: S = CdLines & vbCrLf & EchoLin
FcmdWaitzCdLines = WrtStr(S, T)
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

Sub Exp()
ExpzPj CurPj
End Sub

Sub BrwGitCmitCdLines()
BrwFt GitCmitCdLines("PthA", "Msg", ReInit:=True)
End Sub

Sub BrwGitPushCdLines()
BrwFt GitPushCdLines("A")
End Sub
Private Function PjNm$(CmitgPth)
If Fdr(ParPth(CmitgPth)) <> ".source" Then Thw CSub, "Not source path", "CmitgPth", CmitgPth
PjNm = Fdr(CmitgPth)
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

