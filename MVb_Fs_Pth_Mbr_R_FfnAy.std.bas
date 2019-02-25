Attribute VB_Name = "MVb_Fs_Pth_Mbr_R_FfnAy"
Option Explicit
Private O$(), A_Spec$, A_Atr As FileAttribute ' Used in PthPthAyR/FFnAyR

Function EmpPthAyR(Pth) As String()
Dim I
For Each I In Itr(SubPthAyR(Pth))
    If IsEmpPth(I) Then PushI EmpPthAyR, I
Next
End Function

Function EntAyR(Pth, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = FilSpec
A_Atr = Atr
EntAyR1 Pth
EntAyR = O
End Function

Private Sub EntAyR1(Pth)
Ass HasPth(Pth)
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPthAyR1: (Each 1000): " & Pth
PushI O, Pth
PushIAy O, FfnAy(Pth, A_Spec, A_Atr)
Dim I, P$()
P = SubPthAyz(Pth, A_Spec, A_Atr)
For Each I In Itr(P)
    EntAyR1 I
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
    Act = FfnAyR(Pth, Spec, Atr)
    Brw Act
    Stop
    Return
End Sub
Function FfnAyR(Pth, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
FfnAyR1 Pth
FfnAyR = O
End Function

Private Sub FfnAyR1(Pth)
PushIAy O, FfnAy(Pth, A_Spec, A_Atr)
If Sz(O) Mod 1000 = 0 Then InfoLin CSub, "...Reading", "#Ffn-read", Sz(O)
Dim P$(): P = SubPthAyz(Pth, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Sub
Dim I
For Each I In P
    FfnAyR1 I
Next
End Sub

Private Sub ZZ_EntAyR()
Dim A$(): A = EntAyR("C:\users\user\documents\")
Debug.Print Sz(A)
Stop
DmpAy A
End Sub

Private Sub Z_EmpPthAyR()
Brw EmpPthAyR(TmpRoot)
End Sub

Private Sub Z_EntAy()
BrwPth EntAy(TmpRoot)
End Sub

Private Sub Z_RmvEmpPthR()
RmvEmpPthR TmpRoot
End Sub

Private Sub Z()
'EmpPthAyR
'EntAyR
'FFnAyR
'PthPthAyR
End Sub

