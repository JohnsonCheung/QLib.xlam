Attribute VB_Name = "QVb_Fs_Pth_Mbr_R_FfnAy"
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Mbr_R_Ffny."
Private Const Asm$ = "QVb"
Private O$(), A_Spec$ ' Used in PthPthSyR/FFnAyR

Function EmpPthSyR(Pth) As String()
Dim I
For Each I In Itr(SubPthSyR(Pth))
    If IsPthOfEmp(I) Then PushI EmpPthSyR, I
Next
End Function

Function EntAyR(Pth, Optional FilSpec$ = "*.*") As String()
Erase O
A_Spec = FilSpec
EntAyR1 Pth
EntAyR = O
Erase O
End Function

Private Sub EntAyR1(Pth)
Ass HasPth(Pth)
If Si(O) Mod 1000 = 0 Then Debug.Print "EntAyR1: (Each 1000): " & Pth
PushI O, Pth
PushIAy O, Ffny(Pth, A_Spec)
Dim I, P$()
'P = SubPthSyR(Pth, A_Spec)
For Each I In Itr(P)
    EntAyR1 I
Next
End Sub
Private Sub Z_FfnyR()
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
    Act = FfnyR(Pth, Spec)
    Brw Act
    Stop
    Return
End Sub
Function FfnyR(Pth, Optional Spec$ = "*.*") As String()
Erase O
A_Spec = Spec
FfnyR1 Pth
FfnyR = O
End Function

Private Sub FfnyR1(Pth)
PushIAy O, Ffny(Pth, A_Spec)
If Si(O) Mod 1000 = 0 Then InfLin CSub, "...Reading", "#Ffn-read", Si(O)
Dim P$(): P = SubPthy(Pth)
If Si(P) = 0 Then Exit Sub
Dim I
For Each I In P
    FfnyR1 I
Next
End Sub

Private Sub ZZ_EntAyR()
Dim A$(): A = EntAyR("C:\users\user\documents\")
Debug.Print Si(A)
Stop
DmpAy A
End Sub

Private Sub Z_EmpPthSyR()
Brw EmpPthSyR(TmpRoot)
End Sub

Private Sub Z_EntAy()
BrwPth EntAy(TmpRoot)
End Sub

Private Sub Z_RmvEmpPthR()
RmvEmpPthR TmpRoot
End Sub

Private Sub ZZ()
'EmpPthSyR
'EntAyR
'FFnAyR
'PthPthSyR
End Sub

