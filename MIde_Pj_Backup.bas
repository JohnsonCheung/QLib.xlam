Attribute VB_Name = "MIde_Pj_Backup"
Option Explicit
Function PjfBackup$()
PjfBackup = PjfBackupzPj(CurPj)
End Function

Function PjfBackupzPj$(A As VBProject)
PjfBackupzPj = FfnBackup(Pjf(A))
End Function
