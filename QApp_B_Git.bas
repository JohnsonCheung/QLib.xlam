Attribute VB_Name = "QApp_B_Git"
Option Compare Text
Option Explicit
Private Const Asm$ = "QApp"
Private Const CMod$ = "MApp_Git."
':Fwcmd$ = "It is a TmpFfn and the content is given-CmdLines plus EchoLin which is creating a Fcmdw."
':WaitgFfn$ = "Fcmdw is a temp file without any content.  It is created at end of the Fwcmd."

Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Dim L$ 'CmdLines
    Erase XX
    Dim P$: P = SrcpP
    Dim Pjn$: Pjn = PjnzSrcp(P)
    X "Cd """ & P & """"
    If ReInit Then X "Rd .git /s/q"
    X "git init"        'If already init, it will do nothing
    X "git add -A"
    X FmtQQ("git commit -m ""?""", Msg)
    X "Pause"
    L = JnCrLf(XX)
    Erase XX

Dim F$
    F = TmpCmd("Cmit")
    WrtStr L, F
Shell F, vbMaximizedFocus
End Sub

Sub GitPush()
Dim L$
    Dim P$: P = SrcpP
    Erase XX
    X FmtQQ("Cd ""?""", P)
    X FmtQQ("git push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", PjnzSrcp(P))
    X "Pause ....."
    L = JnCrLf(XX)
Dim F$
    F = TmpCmd("Push")
    WrtStr L, F
Shell F, vbMaximizedFocus
End Sub

Function HasInternet() As Boolean
Stop
End Function

Private Function PjnzSrcp$(Srcp$)
If Fdr(ParPth(Srcp)) <> ".Src" Then Thw CSub, "Not source path", "CmitgPth", Srcp
PjnzSrcp = Fdr(Srcp)
End Function

Private Sub Z()
MApp_Commit:
End Sub

Function RunFps1&(Fps1$, Optional PmStr$)
RunFps1 = RunFcmd("PowerShell", FmtQQ("""?""", Fps1) & PpdSpcIf(PmStr))
End Function

