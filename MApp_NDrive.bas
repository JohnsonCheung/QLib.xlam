Attribute VB_Name = "MApp_NDrive"
Option Explicit


Sub MapNDrive()
RmvNDrive
Shell "Subst N: c:\users\user\desktop\MHD"
End Sub

Sub RmvNDrive()
Shell "Subst /d N:"
End Sub
