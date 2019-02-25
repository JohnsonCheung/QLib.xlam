Attribute VB_Name = "MIde_Pj_Dte"
Option Explicit
Function PjDteFb(A) As Date
Static Y As New Access.Application
Y.OpenCurrentDatabase A
Y.Visible = False
PjDteFb = AcsPjDte(Y)
Y.CloseCurrentDatabase
End Function

Function PjDtePjf(Pjf) As Date
Select Case True
Case IsFxa(Pjf): PjDtePjf = FfnDte(Pjf)
Case IsFb(Pjf): PjDtePjf = PjDteFb(Pjf)
Case Else: Stop
End Select
End Function

Function AcsPjDte(A As Access.Application)
Dim O As Date
Dim M As Date
M = MaxItrPrp(A.CurrentProject.AllForms, "DateModified")
O = Max(O, M)
O = Max(O, MaxItrPrp(A.CurrentProject.AllModules, "DateModified"))
O = Max(O, MaxItrPrp(A.CurrentProject.AllReports, "DateModified"))
AcsPjDte = O
End Function
