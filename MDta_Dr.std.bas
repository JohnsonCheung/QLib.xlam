Attribute VB_Name = "MDta_Dr"
Option Explicit
Function DrTLinzbTyAy(TLin, VbTyAy() As VbVarType) As Variant()

End Function
Function VbTyShtTy(ShtTy$) As VbVarType
Dim O As VbVarType
Select Case ShtTy
Case ""
Case ""
End Select
End Function
Function VbTyAyShtTyLis(ShtTyLis$) As VbVarType()
Dim J%
For J = 1 To Len(ShtTyLis)
    PushI VbTyAyShtTyLis, VbTyShtTy(Mid(ShtTyLis, J, 1))
Next
End Function
Function DrTLinShtTyLis(TLin, ShtTyLis$) As Variant()
DrTLinShtTyLis = DrTLinzbTyAy(TLin, VbTyAyShtTyLis(ShtTyLis))
End Function

Function DrzDrs(A As Drs, Optional CC, Optional Row&)
DrzDrs = AywIxAy(A.Dry()(Row), IxAy(A.Fny, FnyzFF(CC)))
End Function
