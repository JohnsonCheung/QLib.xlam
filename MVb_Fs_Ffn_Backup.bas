Attribute VB_Name = "MVb_Fs_Ffn_Backup"
Option Explicit
Sub BrkBkPth()
BrwPth BkPthzPj(CurPj)
End Sub
Function BkPthzPj$(A As VBProject)
BkPthzPj = BkPth(Pjf(A))
End Function

Function BkPth$(Ffn$)
BkPth = AddFdrApEns(Pth(Ffn$), ".Backup", Fn(Ffn$))
End Function
Function LasBkFfn$(Ffn$)
Dim BkPth1$: BkPth1 = BkPth(Ffn$)
'LasInstFdr (BkPth1)
End Function
Function BkFfn$(Ffn$)
BkFfn = AddFdrEns(BkPth(Ffn$), TmpNm)
End Function
Function BackupFfn$(Ffn$)
BackupFfn = CpyFfnzToPth(Ffn$, BkFfn(Ffn$))
End Function

Sub RplFfn(Ffn$, ByFfn$)
BackupFfn Ffn
If DltFfnDone(Ffn$) Then
    Name Ffn As ByFfn
End If
End Sub
