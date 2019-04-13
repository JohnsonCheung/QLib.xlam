Attribute VB_Name = "MVb_Fs_Pth_Mbr_R_SubPthAy"
Option Explicit
Private xx$()
Function SubPthAyR(Pth) As String()
Erase xx
SubPthAyRz Pth
SubPthAyR = xx
Erase xx
End Function
Private Sub SubPthAyRz(Pth)
Dim O$(), P
O = SubPthAy(Pth)
PushIAy xx, O
For Each P In Itr(O)
    SubPthAyRz P
Next
End Sub
