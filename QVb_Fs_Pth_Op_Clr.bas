Attribute VB_Name = "QVb_Fs_Pth_Op_Clr"
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Op_Clr."
Private Const Asm$ = "QVb"

Sub ClrPth(Pth)
DltFfnyAyIf Ffny(Pth)
End Sub

Private Sub Z_ClrPthFil()
ClrPthFil TmpRoot
End Sub

Sub ClrPthFil(Pth)
If Not IsPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(Ffny(Pth))
   DltFfn F
Next
End Sub

