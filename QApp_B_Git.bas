Attribute VB_Name = "QApp_B_Git"
Option Compare Text
Option Explicit
Private Const Asm$ = "QApp"
Private Const CMod$ = "MApp_Git."
':Fwcmd$ = "It is a TmpFfn and the content is given-CmdLines plus EchoLin which is creating a Fcmdw."
':WaitgFfn$ = "Fcmdw is a temp file without any content.  It is created at end of the Fwcmd."

Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
Dim X As New Bfr
    Dim P$:     P = SrcpP
    X.Var "Cd """ & P & """"
    If ReInit Then X "Rd .git /s/q"
    X.Var "git init"        'If already init, it will do nothing
    X.Var "git add -A"
    X.Var FmtQQ("git commit -m ""?""", Msg)
    X.Var "Pause"
Dim F$
    Dim Pjn$: Pjn = PjnzSrcp(P)
                F = TmpFcmd("Cmit")
WrtStr X.Lines, F
Shell F, vbMaximizedFocus
End Sub

Sub GitPush()
Dim X As New Bfr
    Dim P$: P = SrcpP
    X.Var FmtQQ("Cd ""?""", P)
    X.Var FmtQQ("git push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", PjnzSrcp(P))
    X.Var "Pause ....."
Dim F$
    F = TmpFcmd("Push")
    WrtStr X.Lines, F
Shell F, vbMaximizedFocus
End Sub

Function HasInternet() As Boolean
Stop
End Function

Function PjnzSrcp$(Srcp$)
If Fdr(ParPth(Srcp)) <> ".Src" Then Thw CSub, "Not source path", "CmitgPth", Srcp
PjnzSrcp = Fdr(Srcp)
End Function

Private Sub Z()
MApp_Commit:
End Sub

