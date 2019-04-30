Attribute VB_Name = "MIde_Gen_Pjf_Expg"
Option Explicit

Function SrcpSyOfExpgInst() As String()
Dim P$, I
For Each I In Itr(SubPthSyR(ExpgPth))
    P = I
    If IsInstScrp(P) Then
        PushI SrcpSyOfExpgInst, P
    End If
Next
End Function

Function SrcpSyOfExpgInstWoNonEmpDist() As String()
Dim Pth$, I, Dist$
For Each I In Itr(SrcpSyOfExpgInst)
    Pth = I
    Dist = SiblingPth(Pth, "Dist")
    Select Case True
    Case Not IsPth(Dist), IsEmpPth(Dist): PushI SrcpSyOfExpgInstWoNonEmpDist, Pth
    End Select
Next
End Function

Sub GenExpg()
Dim Ay$(): Ay = SrcpSyOfExpgInstWoNonEmpDist
If Si(Ay) = 0 Then Exit Sub
Dim Srcp$, I, Xls As Excel.Application, Acs As Access.Application
Set Xls = NewXls: Set Acs = NewAcs
For Each I In Itr(Ay)
    Srcp = I
    Stamp "GenExpg: Begin"
    Stamp "GenExpg: Srcp " & Srcp
    CrtDistFxa Srcp, Xls
    GenFba Srcp, Acs
    Stamp "GenExpg: End"
Next
AcsQuit Acs
QuitXls Xls
End Sub

Private Sub Z_SrcpSyOfExpgInst()
DmpAy SrcpSyOfExpgInst
End Sub

Private Sub Z_SrcpSyOfExpgInstWoNonEmpDist()
DmpAy SrcpSyOfExpgInstWoNonEmpDist
End Sub
