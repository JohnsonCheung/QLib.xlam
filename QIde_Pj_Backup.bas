Attribute VB_Name = "QIde_Pj_Backup"
Option Explicit
Private Const CMod$ = "MIde_Pj_Backup."
Private Const Asm$ = "QIde"
Sub BackupPj()
BackupFfn Pjf(CurPj)
End Sub
