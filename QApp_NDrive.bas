Attribute VB_Name = "QApp_NDrive"
Option Explicit
Private Const CMod$ = "MApp_NDrive."
Private Const Asm$ = "QApp"
Sub MapNDrive()
RmvNDrive
Shell "Subst N: c:\users\user\desktop\MHD"
End Sub

Sub RmvNDrive()
Shell "Subst /d N:"
End Sub