Attribute VB_Name = "MVb_Fs_Pth_Mbr_R_SubPthAy"
Option Explicit
Private XX$()
Function SubPthAyR(Pth) As String()
Erase XX
SubPthAyRz Pth
SubPthAyR = XX
Erase XX
End Function
Private Sub SubPthAyRz(Pth)
Dim O$(), P
O = SubPthAy(Pth)
PushIAy XX, O
For Each P In Itr(O)
    SubPthAyRz P
Next
End Sub
