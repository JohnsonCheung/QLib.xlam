Attribute VB_Name = "QVb_Zip"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Zip."
Private Const Asm$ = "QVb"
Sub ZipPth(Pth, Optional PthKd$ = "Path")
ThwIf_PthNotExist Pth, CSub
End Sub
