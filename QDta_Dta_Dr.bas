Attribute VB_Name = "QDta_Dta_Dr"
Option Explicit
Private Const CMod$ = "MDta_Dr."
Private Const Asm$ = "QDta"
Function DrzTLinVbTyAy(TLin, VbTyAy() As VbVarType) As Variant()

End Function
Function VbTyzShtTy(ShtTy$) As VbVarType
Dim O As VbVarType
Select Case ShtTy
Case ""
Case ""
End Select
End Function
Function VbTyAyzShtTyLis(ShtTyLis$) As VbVarType()
Dim J%
For J = 1 To Len(ShtTyLis)
    PushI VbTyAyzShtTyLis, VbTyzShtTy(Mid(ShtTyLis, J, 1))
Next
End Function
Function DrzTLinShtTyLis(TLin, ShtTyLis$) As Variant()
DrzTLinShtTyLis = DrzTLinVbTyAy(TLin, VbTyAyzShtTyLis(ShtTyLis))
End Function

Function DrvzDrs(A As Drs, Cny$(), Optional Row&)
DrvzDrs = AywIxy(A.Dry()(Row), Ixy(A.Fny, Cny))
End Function

