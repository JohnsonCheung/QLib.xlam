Attribute VB_Name = "QVb_Fs_Ffn_Backup"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Backup."
Private Const Asm$ = "QVb"
Sub BrwBkPth()
BrwPth BkPthzPj(CurPj)
End Sub
Function BkPthzPj$(A As VBProject)
BkPthzPj = BkPth(Pjf(A))
End Function
Function BkRoot$(Pth$)
BkRoot = AddFdrEns(Pth, ".Backup")
End Function
Function BkHom$(Ffn$)
BkHom = AddFdrEns(BkRoot(Pth(Ffn)), Fn(Ffn))
End Function

Function LasBkFfnC$()
LasBkFfnC = LasBkFfn(PjfC)
End Function
Function LasBkFfn$(Ffn$)
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrSyzIsInst(H)
Dim Fdr$: Fdr = MaxAy(F)
LasBkFfn = H & Fdr & "\" & Fn(Ffn)
End Function
Function BkPth$(Ffn$)
BkPth = AddFdr(BkHom(Ffn), TmpNm)
End Function

Function BkFfn$(Ffn$)
BkFfn = BkPth(Ffn) & Fn(Ffn)
End Function
Function BackupFfn$(Ffn$)
Dim T$: T = BkFfn(Ffn)
EnsPthzAllSeg Pth(T)
CpyFfn Ffn, T
BackupFfn = T
End Function

Sub RplFfn(Ffn$, ByFfn$)
BackupFfn Ffn
If DltFfnDone(Ffn$) Then
    Name Ffn As ByFfn
End If
End Sub
