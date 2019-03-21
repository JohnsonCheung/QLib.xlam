Attribute VB_Name = "MIde_Gen_Pjf_Expg"
Option Explicit
Sub Z2()
GenExpg
End Sub

Sub GenExpg()
Dim Ay$(): Ay = SrcpAyzExpgzInstzNoNonEmpDist
If Si(Ay) = 0 Then Exit Sub
Dim Srcp, Xls As Excel.Application, Acs As Access.Application
Set Xls = NewXls: Set Acs = NewAcs
For Each Srcp In Itr(Ay)
    Stamp "GenExpg: Begin"
    Stamp "GenExpg: Srcp " & Srcp
    DistFxazSrcp Srcp, Xls
    GenFba Srcp, Acs
    Stamp "GenExpg: End"
Next
AcsQuit Acs
XlsQuit Xls
End Sub

Function SrcpAyzExpgzInst() As String()
Dim P
For Each P In Itr(SubPthAyR(ExpgPth))
    If IsSrcpInst(P) Then
        PushI SrcpAyzExpgzInst, P
    End If
Next
End Function
Private Sub Z_SrcpAyzExpgzInstzNoNonEmpDist()
DmpAy SrcpAyzExpgzInstzNoNonEmpDist
End Sub
Private Sub Z_SrcpAyzExpgzInst()
DmpAy SrcpAyzExpgzInst
End Sub
Function SrcpAyzExpgzInstzNoNonEmpDist() As String()
Dim Pth, Dist$
For Each Pth In Itr(SrcpAyzExpgzInst)
    Dist = SiblingPth(Pth, "Dist")
    Select Case True
    Case Not IsPth(Dist), IsEmpPth(Dist): PushI SrcpAyzExpgzInstzNoNonEmpDist, Pth
    End Select
Next
End Function

