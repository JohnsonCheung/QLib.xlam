Attribute VB_Name = "QIde_Mth_Lin_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is."
Private Const Asm$ = "QIde"
Function IsMthLin(Lin) As Boolean
IsMthLin = MthKd(Lin) <> ""
End Function
Function IsMthLinzNm(Lin, Nm) As Boolean
IsMthLinzNm = Mthn(Lin) = Nm
End Function

