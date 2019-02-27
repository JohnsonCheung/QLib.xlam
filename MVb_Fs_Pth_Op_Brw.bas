Attribute VB_Name = "MVb_Fs_Pth_Op_Brw"
Option Explicit

Sub BrwPthVC(Pth)
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
End Sub

Sub BrwPth(Pth)
Shell FmtQQ("Explorer ""?""", Pth)
End Sub

