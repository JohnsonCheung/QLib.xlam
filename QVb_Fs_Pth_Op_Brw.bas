Attribute VB_Name = "QVb_Fs_Pth_Op_Brw"
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Op_Brw."
Private Const Asm$ = "QVb"

Sub VcPth(Pth$)
If Not HasPth(Pth) Then Debug.Print "No such path": Exit Sub
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
End Sub

Sub BrwPth(Pth$)
If Not HasPth(Pth) Then Debug.Print "No such path": Exit Sub
Shell FmtQQ("Explorer ""?""", Pth)
End Sub

