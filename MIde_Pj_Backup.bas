Attribute VB_Name = "MIde_Pj_Backup"
Option Explicit
Sub BackupPj()
PjfBackupzPj CurPj
End Sub

Function PjfBackupzPj$(A As VBProject)
PjfBackupzPj = FfnBackup(Pjf(A))
End Function
