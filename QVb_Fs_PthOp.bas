Attribute VB_Name = "QVb_Fs_PthOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Op_Brw."
Private Const Asm$ = "QVb"

Sub VcPth(Pth)
If Not HasPth(Pth) Then Debug.Print "No such path": Exit Sub
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
End Sub

Sub BrwPth(Pth)
If Not HasPth(Pth) Then Debug.Print "No such path": Exit Sub
Shell FmtQQ("Explorer ""?""", Pth)
End Sub


Sub ClrPth(Pth)
DltFfnyAyIf Ffny(Pth)
End Sub

Private Sub Z_ClrPthFil()
ClrPthFil TmpRoot
End Sub

Sub ClrPthFil(Pth)
If Not IsPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(Ffny(Pth))
   DltFfn F
Next
End Sub


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

Sub RenPthAddPfx(Pth, Pfx)
RenPth Pth, AddPfxzPth(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
If HasPth(NewPth) Then Thw CSub, "NewPth Has", "Pth NewPth", Pth, NewPth
If Not HasPth(Pth) Then Thw CSub, "Pth not Has", "Pth NewPth", Pth, NewPth
Fso.GetFolder(Pth).Name = NewPth
End Sub



Private Sub Z_RmvEmpSubDir()
RmvEmpSubDir TmpPth
End Sub



'
