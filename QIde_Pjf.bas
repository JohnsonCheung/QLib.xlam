Attribute VB_Name = "QIde_Pjf"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Pjf."
Public PjfAcs As New Access.Application
Public PjfXls As New Excel.Application

Sub ClsPjf(Pjf)
Const CSub$ = CMod & "ClsPjf"
Select Case True
Case IsFxa(Pjf): RmvPjzXlsPjf PjfXls, Pjf
Case IsFb(Pjf):  ClsDbzAcs PjfAcs
Case Else: Thw CSub, "Invalid Pjf, should be Fxa or Fb", "Pjf", Pjf
End Select
End Sub
Function VbezFba(Fba) As Vbe
'OpnPjf Pjf

End Function
Function VbezFxa(Fxa) As Vbe
'OpnPjf Pjf

End Function
Function VbezPjf(Pjf) As Vbe
Const CSub$ = CMod & "VbezPjf"
Select Case True
Case IsFxa(Pjf): Set VbezPjf = VbezFba(Pjf)
Case IsFba(Pjf):  Set VbezPjf = VbezFxa(Pjf)
Case Else: Thw CSub, "Invalid Pjf, should be Fxa or Fba", "Pjf", Pjf
End Select
End Function
Sub OpnPjf(Pjf)  ' Return either Exl.Application (Exl) or Acs.Application (Function-static)
Select Case True
Case IsFxa(Pjf): PjfXls.Workbooks.Open Pjf
Case IsFba(Pjf):  OpnFb PjfAcs, Pjf
Case Else: Stop
End Select
End Sub

Sub RmvPjzXlsPjf(Exl As Excel.Application, Pjf)
Dim Pj As VBProject
Set Pj = PjzPjf(Exl.Vbe, Pjf)
Pj.Collection.Remove Pj
End Sub

Function TmpFxa$(Optional Fdr$, Optional Fnn$)
TmpFxa = TmpFfn(".xlam", Fdr, Fnn)
End Function

Private Sub Z()
QIde_Pjf:
End Sub
