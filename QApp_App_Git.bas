Attribute VB_Name = "QApp_App_Git"
Option Explicit
Private Const Asm$ = "QApp"
Private Const CMod$ = "MApp_Git."
Public Const DoczFwcmd$ = "It is a TmpFfn and the content is given-CmdLines plus EchoLin which is creating a Fcmdw."
Public Const DoczWaitgFfn$ = "Fcmdw is a temp file without any content.  It is created at end of the Fwcmd."

Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Dim mCmdLines$: mCmdLines = CmdLineszCmitg(SrcpP, Msg, ReInit)
Dim mFcmdw$: mFcmdw = Fcmdw(mCmdLines)
WaitFcmdw mFcmdw, DftWaitOpt
End Sub

Function FcmdwzPushg$()
FcmdwzPushg = Fcmdw(CmdLineszPushg(Srcp(CPj)))
End Function

Sub GitPush()
WaitFcmdw FcmdwzPushg, DftWaitOpt
End Sub

Private Function CmdLineszCmitg$(CmitgPth$, Msg$, ReInit As Boolean)
Erase XX
Dim Pj$: Pj = PjnzCmitgPth(CmitgPth)
X "Cd """ & CmitgPth & """"
If ReInit Then X "Rd .git /s/q"
X "git init"        'If already init, it will do nothing
X "git add -A"
X FmtQQ("git commit -m ""?""", Msg)
X "Pause"
CmdLineszCmitg = JnCrLf(XX)
Erase XX
End Function


Function HasInternet() As Boolean
Stop
End Function

Private Function CmdLineszPushg$(CmitgPth$)
Dim O$(), Cd$, GitPush, T
Push O, FmtQQ("Cd ""?""", CmitgPth)
Push O, FmtQQ("git push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", PjnzCmitgPth(CmitgPth))
Push O, "Pause"
CmdLineszPushg = JnCrLf(O)
End Function

Sub BrwCmdLineszCmitg()
BrwFt CmdLineszCmitg("PthA", "Msg", ReInit:=True)
End Sub

Sub BrwCmdLineszPushg()
BrwFt CmdLineszPushg("A")
End Sub
Private Function PjnzCmitgPth$(CmitgPth$)
Const CSub$ = CMod & "Pjn"
If Fdr(ParPth(CmitgPth)) <> ".Src" Then Thw CSub, "Not source path", "CmitgPth", CmitgPth
PjnzCmitgPth = Fdr(CmitgPth)
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

Private Sub ZZ()
MApp_Commit:
End Sub

Function RunFps1&(Fps1$, Optional PmStr$)
RunFps1 = RunFcmd("PowerShell", FmtQQ("""?""", Fps1) & PpdSpcIf(PmStr))
End Function

