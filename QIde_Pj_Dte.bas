Attribute VB_Name = "QIde_Pj_Dte"
Option Explicit
Private Const CMod$ = "MIde_Pj_Dte."
Private Const Asm$ = "QIde"
Function PjDtezFb(Fb$) As Date
Static Y As New Access.Application
Y.OpenCurrentDatabase Fb
Y.Visible = False
PjDtezFb = PjDtezAcs(Y)
Y.CloseCurrentDatabase
End Function

Function PjDtezPjf(Pjf$) As Date
Select Case True
Case IsFxa(Pjf): PjDtezPjf = DtezFfn(Pjf)
Case IsFb(Pjf): PjDtezPjf = PjDtezFb(Pjf)
Case Else: Stop
End Select
End Function

Function PjDtezAcs(A As Access.Application)
Dim O As Date
Dim M As Date
M = MaxzItrPrp(A.CurrentProject.AllForms, "DateModified")
O = Max(O, M)
O = Max(O, MaxzItrPrp(A.CurrentProject.AllModules, "DateModified"))
O = Max(O, MaxzItrPrp(A.CurrentProject.AllReports, "DateModified"))
PjDtezAcs = O
End Function
