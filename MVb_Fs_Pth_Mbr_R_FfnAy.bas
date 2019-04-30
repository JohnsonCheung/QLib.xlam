Attribute VB_Name = "MVb_Fs_Pth_Mbr_R_FfnAy"
Option Explicit
Private O$(), A_Spec$ ' Used in PthPthSyR/FFnAyR

Function EmpPthSyR(Pth$) As String()
Dim I
For Each I In Itr(SubPthSyR(Pth))
    If IsEmpPth(I) Then PushI EmpPthSyR, I
Next
End Function

Function EntSyR(Pth$, Optional FilSpec$ = "*.*") As String()
Erase O
A_Spec = FilSpec
EntSyR1 Pth
EntSyR = O
Erase O
End Function

Private Sub EntSyR1(Pth$)
Ass HasPth(Pth)
If Si(O) Mod 1000 = 0 Then Debug.Print "EntSyR1: (Each 1000): " & Pth
PushI O, Pth
PushIAy O, FfnSy(Pth, A_Spec)
Dim I, P$()
P = SubPthSyz(Pth, A_Spec)
For Each I In Itr(P)
    EntSyR1 I
Next
End Sub
Private Sub Z_FfnAyR()
Dim Pth, Spec$, Atr As FileAttribute
GoSub T0
GoSub T1
Exit Sub
T0:
    Pth = "C:\Users\User\Documents\Projects\Vba"
    GoTo Tst
T1:
    Pth = "C:\Users\User\Documents\WindowsPowershell\"
    GoTo Tst
Tst:
    Act = FfnAyR(Pth, Spec)
    Brw Act
    Stop
    Return
End Sub
Function FfnAyR(Pth$, Optional Spec$ = "*.*") As String()
Erase O
A_Spec = Spec
FfnAyR1 Pth
FfnAyR = O
End Function

Private Sub FfnAyR1(Pth$)
PushIAy O, FfnSy(Pth, A_Spec)
If Si(O) Mod 1000 = 0 Then InfLin CSub, "...Reading", "#Ffn-read", Si(O)
Dim P$(): P = SubPthSyz(Pth, A_Spec)
If Si(P) = 0 Then Exit Sub
Dim I
For Each I In P
    FfnAyR1 I
Next
End Sub

Private Sub ZZ_EntSyR()
Dim A$(): A = EntSyR("C:\users\user\documents\")
Debug.Print Si(A)
Stop
DmpAy A
End Sub

Private Sub Z_EmpPthSyR()
Brw EmpPthSyR(TmpRoot)
End Sub

Private Sub Z_EntSy()
BrwPth EntSy(TmpRoot)
End Sub

Private Sub Z_RmvEmpPthR()
RmvEmpPthR TmpRoot
End Sub

Private Sub Z()
'EmpPthSyR
'EntSyR
'FFnAyR
'PthPthSyR
End Sub

