Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Option Explicit
Function FfnBackup$(Ffn)
FfnBackup = CpyFilzToPth(Ffn, AddFdrEns(Pth(Ffn), ".FfnBackup", Fn(Ffn), TmpNm))
End Function
