Attribute VB_Name = "MIde_Pj_Backup"
Option Explicit
Function BackupPj$()
BackupPj = BackupzPj(CurPj)
End Function

Function BackupzPj$(A As VBProject)
BackupzPj = Backup(Pjf(A))
End Function

