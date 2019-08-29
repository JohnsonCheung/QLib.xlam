Attribute VB_Name = "QApp_B_MapDrive"
Option Compare Text
Option Explicit
Private Const CMod$ = "MApp_NDrive."
Private Const Asm$ = "QApp"

Sub MapDrive(Drv$, Pth$)
RmvDrive Drv
Shell FmtQQ("Subst ? ""?""", Drv, Pth)
End Sub

Sub MapNDrive()
MapDrive "N:", "c:\users\user\desktop\MHD"
End Sub

Sub RmvDrive(Drv$)
Shell "Subst /d " & Drv
End Sub

Sub RmvNDrive()
RmvDrive "N:"
End Sub
