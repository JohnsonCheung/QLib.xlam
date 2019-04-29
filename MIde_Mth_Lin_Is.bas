Attribute VB_Name = "MIde_Mth_Lin_Is"
Option Explicit
Function IsMthLin(Lin$) As Boolean
IsMthLin = MthKd(Lin$) <> ""
End Function
Function IsMthLinzNm(Lin$, Nm) As Boolean
IsMthLinzNm = MthNm(Lin) = Nm
End Function

