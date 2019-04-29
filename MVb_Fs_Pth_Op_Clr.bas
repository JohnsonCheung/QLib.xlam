Attribute VB_Name = "MVb_Fs_Pth_Op_Clr"
Option Explicit

Sub ClrPth(Pth$)
DltFfnyAyIf FfnSy(Pth)
End Sub

Private Sub Z_ClrPthFil()
ClrPthFil TmpRoot
End Sub

Sub ClrPthFil(Pth$)
If Not IsPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(FfnSy(Pth))
   DltFfn CStr(F)
Next
End Sub

