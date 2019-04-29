Attribute VB_Name = "MIde_Gen_Pjf_Expg"
Option Explicit

Function ScrpAyOfExpgInst() As String()
Dim P
For Each P In Itr(SubPthAyR(ExpgPth))
    If IsInstScrp(P) Then
        PushI ScrpAyOfExpgInst, P
    End If
Next
End Function

Function SrcpAyOfExpgInstWoNonEmpDist() As String()
Dim Pth, Dist$
For Each Pth In Itr(ScrpAyOfExpgInst)
    Dist = SiblingPth(Pth, "Dist")
    Select Case True
    Case Not IsPth(Dist), IsEmpPth(Dist): PushI SrcpAyOfExpgInstWoNonEmpDist, Pth
    End Select
Next
End Function

Sub GenExpg()
Dim Ay$(): Ay = SrcpAyOfExpgInstWoNonEmpDist
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
QuitXls Xls
End Sub

Private Sub Z_ScrpAyOfExpgInst()
DmpAy ScrpAyOfExpgInst
End Sub

Private Sub Z_SrcpAyOfExpgInstWoNonEmpDist()
DmpAy SrcpAyOfExpgInstWoNonEmpDist
End Sub
