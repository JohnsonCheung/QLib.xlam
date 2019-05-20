Attribute VB_Name = "QVb_Fs_Pth_Op_Rmv"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Op_Rmv."
Private Const Asm$ = "QVb"
Private Sub Z_RmvEmpPthR()
Debug.Print "Before-----"
D EmpPthSyR(TmpRoot)
RmvEmpPthR TmpRoot
Debug.Print "After-----"
D EmpPthSyR(TmpRoot)
End Sub
Sub RmvEmpPthR(Pth)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Ay = EmpPthSyR(Pth): If Si(Ay) = 0 Then Exit Sub
    For Each I In Ay
        RmDir I
    Next
    GoTo Lp
End Sub

Sub RmvEmpSubDir(Pth)
Dim SubPth
For Each SubPth In Itr(SubPthy(Pth))
   RmvPthIfEmp SubPth
Next
End Sub

Sub RmvPthIfEmp(Pth)
If IsPthOfEmp(Pth) Then Exit Sub
RmDir Pth
End Sub



Private Sub ZZ_RmvEmpSubDir()
RmvEmpSubDir TmpPth
End Sub

