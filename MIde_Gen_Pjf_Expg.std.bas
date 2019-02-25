Attribute VB_Name = "MIde_Gen_Pjf_Expg"
Option Explicit
Sub Z2()
GenExpg
End Sub

Sub GenExpg()
Dim Ay$(): Ay = SrcPthAyzExpgzInstzNoNonEmpDist
If Sz(Ay) = 0 Then Exit Sub
Dim SrcPth, Xls As Excel.Application, Acs As Access.Application
Set Xls = NewXls: Set Acs = NewAcs
For Each SrcPth In Itr(Ay)
    Stamp "GenExpg: Begin"
    Stamp "GenExpg: SrcPth " & SrcPth
    GenFxa SrcPth, Xls
    GenFba SrcPth, Acs
    Stamp "GenExpg: End"
Next
AcsQuit Acs
XlsQuit Xls
End Sub

Function SrcPthAyzExpgzInst() As String()
Dim P
For Each P In Itr(SubPthAyR(ExpgPth))
    If IsSrcPthInst(P) Then
        PushI SrcPthAyzExpgzInst, P
    End If
Next
End Function
Private Sub Z_SrcPthAyzExpgzInstzNoNonEmpDist()
DmpAy SrcPthAyzExpgzInstzNoNonEmpDist
End Sub
Private Sub Z_SrcPthAyzExpgzInst()
DmpAy SrcPthAyzExpgzInst
End Sub
Function SrcPthAyzExpgzInstzNoNonEmpDist() As String()
Dim Pth, Dist$
For Each Pth In Itr(SrcPthAyzExpgzInst)
    Dist = SiblingPth(Pth, "Dist")
    Select Case True
    Case Not IsPth(Dist), IsEmpPth(Dist): PushI SrcPthAyzExpgzInstzNoNonEmpDist, Pth
    End Select
Next
End Function

