Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Function Backup$(Ffn)
Backup = CpyFilzToPth(Ffn, AddFdrEns(Pth(Ffn), ".Backup", Fn(Ffn), TmpNm))
End Function
