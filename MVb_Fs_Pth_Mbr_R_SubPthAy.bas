Attribute VB_Name = "MVb_Fs_Pth_Mbr_R_SubPthAy"
Option Explicit
Private XX$()
Function SubPthSyR(Pth$) As String()
Erase XX
X Pth
SubPthSyR = XX
Erase XX
End Function
Private Sub X(Pth$)
Dim O$(), P$, I
O = SubPthSy(Pth)
PushIAy XX, O
For Each I In Itr(O)
    P = I
    X P
Next
End Sub
