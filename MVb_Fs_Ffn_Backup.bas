Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Option Explicit

Function FfnBackup$(Ffn)
FfnBackup = CpyFilzToPth(Ffn, AddFdrEns(Pth(Ffn), ".FfnBackup", Fn(Ffn), TmpNm))
End Function

Function FfnRpl$(Ffn, ByFfn)
FfnBackup Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
    FfnRpl = ByFfn
End If
End Function
