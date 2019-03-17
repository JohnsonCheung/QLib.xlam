Attribute VB_Name = "MVb_Fs_Pth_Op_Rmv"
Option Explicit
Private Sub Z_RmvEmpPthR()
Debug.Print "Before-----"
D EmpPthAyR(TmpRoot)
RmvEmpPthR TmpRoot
Debug.Print "After-----"
D EmpPthAyR(TmpRoot)
End Sub
Sub RmvEmpPthR(Pth)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Ay = EmpPthAyR(Pth): If Si(Ay) = 0 Then Exit Sub
    For Each I In Ay
        RmDir I
    Next
    GoTo Lp
End Sub

Sub RmvEmpSubDir(Pth)
Dim SubPth
For Each SubPth In Itr(SubPthAy(Pth))
   RmvPthIfEmp SubPth
Next
End Sub

Sub RmvPthIfEmp(Pth)
If IsEmpPth(Pth) Then Exit Sub
RmDir Pth
End Sub



Private Sub ZZ_RmvEmpSubDir()
RmvEmpSubDir TmpPth
End Sub

