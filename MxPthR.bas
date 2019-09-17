Attribute VB_Name = "MxPthR"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxPthR."
Private O$(), A_Spec$ ' Used in PrpcAyR/FFnAyR
Private XX$()

Function EmpPthSyR(Pth) As String()
Dim I
For Each I In Itr(SubPthSyR(Pth))
    If HasPthOfEmp(I) Then PushI EmpPthSyR, I
Next
End Function

Function EntAyR(Pth, Optional FilSpec$ = "*.*") As String()
Erase O
A_Spec = FilSpec
EntAyR1 Pth
EntAyR = O
Erase O
End Function

Sub EntAyR1(Pth)
If Si(O) Mod 1000 = 0 Then Debug.Print "EntAyR1: (Each 1000): " & Pth
PushI O, Pth
PushIAy O, FfnAy(Pth, A_Spec)
Dim I, P$()
'P = SubPthSyR(Pth, A_Spec)
For Each I In Itr(P)
    EntAyR1 I
Next
End Sub
Sub Z_FfnAyR()
Dim Pth, Spec$, AtR As FileAttribute
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
Function FfnAyR(Pth, Optional Spec$ = "*.*") As String()
Erase O
A_Spec = Spec
FfnAyR1 Pth
FfnAyR = O
End Function

Sub FfnAyR1(Pth)
PushIAy O, FfnAy(Pth, A_Spec)
If Si(O) Mod 1000 = 0 Then InfLin CSub, "...Reading", "#Ffn-read", Si(O)
Dim P$(): P = SubPthy(Pth)
If Si(P) = 0 Then Exit Sub
Dim I: For Each I In P
    FfnAyR1 I
Next
End Sub

Sub Z_EntAyR()
Dim A$(): A = EntAyR("C:\users\user\documents\")
Debug.Print Si(A)
Stop
DmpAy A
End Sub

Sub Z_EmpPthSyR()
Brw EmpPthSyR(TmpRoot)
End Sub

Sub Z_EntAy()
BrwPth EntAy(TmpRoot)
End Sub

Sub Z_RmvEmpPthR()
Z:
    RmvEmpPthR TmpRoot
    Return
Z1:
    Debug.Print "Before-----"
    D EmpPthSyR(TmpRoot)
    RmvEmpPthR TmpRoot
    Debug.Print "After-----"
    D EmpPthSyR(TmpRoot)
    Return
End Sub

Function SubPthSyR(Pth) As String()
Erase XX
X Pth
SubPthSyR = XX
Erase XX
End Function

Private Sub X(Pth)
Dim O$(), P$, I
O = SubPthy(Pth)
PushIAy XX, O
For Each I In Itr(O)
    P = I
    X P
Next
End Sub
