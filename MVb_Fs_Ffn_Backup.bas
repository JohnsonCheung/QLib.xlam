Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Option Explicit
Function Backup$(Ffn)
Backup = CpyFilzToPth(Ffn, AddFdrEns(Pth(Ffn), ".Backup", Fn(Ffn), TmpNm))
End Function
