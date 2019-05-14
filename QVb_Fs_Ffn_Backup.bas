Attribute VB_Name = "QVb_Fs_Ffn_Backup"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Backup."
Private Const Asm$ = "QVb"
Sub BrwBkPth()
BrwPth BkPthzP(CPj)
End Sub
Function BkPthzP$(P As VBProject)
BkPthzP = BkPth(Pjf(P))
End Function
Function BkRoot$(Pth)
BkRoot = AddFdrEns(Pth, ".Backup")
End Function
Function BkHom$(Ffn)
BkHom = AddFdrEns(BkRoot(Pth(Ffn)), Fn(Ffn))
End Function

Function LasBkPjfP$()
LasBkPjfP = LasBkFfn(PjfP)
End Function
Function LasBkFfn$(Ffn)
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fdr$: Fdr = MaxAy(F)
LasBkFfn = H & Fdr & "\" & Fn(Ffn)
End Function
Function BkPth$(Ffn)
BkPth = AddFdr(BkHom(Ffn), TmpNm)
End Function

Function BkFfn$(Ffn)
BkFfn = BkPth(Ffn) & Fn(Ffn)
End Function
Function BackupFfn$(Ffn)
Dim T$: T = BkFfn(Ffn)
EnsPthzAllSeg Pth(T)
CpyFfn Ffn, T
BackupFfn = T
End Function

Sub RplFfn(Ffn, ByFfn$)
BackupFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
End If
End Sub
