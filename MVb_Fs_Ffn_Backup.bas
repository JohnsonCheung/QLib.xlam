Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Option Explicit
Sub BrkBkPth()
BrwPth BkPthzPj(CurPj)
End Sub
Function BkPthzPj$(A As VBProject)
BkPthzPj = BkPth(Pjf(A))
End Function

Function BkPth$(Ffn)
BkPth = AddFdrEns(Pth(Ffn), ".Backup", Fn(Ffn))
End Function
Function LasBkFfn$(Ffn)
Dim BkPth1$: BkPth1 = BkPth(Ffn)
'LasInstFdr (BkPth1)
End Function
Function BkFfn$(Ffn)
BkFfn = AddFdrEns(BkPth(Ffn), TmpNm)
End Function
Function BackupFfn$(Ffn)
BackupFfn = CpyFilzToPth(Ffn, BkFfn(Ffn))
End Function

Function FfnRpl$(Ffn, ByFfn)
BackupFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
    FfnRpl = ByFfn
End If
End Function
