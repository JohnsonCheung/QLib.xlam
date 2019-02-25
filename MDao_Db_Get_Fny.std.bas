Attribute VB_Name = "MDao_Db_Get_Fny"
Option Explicit

Function DryzDbq(A As Database, Q) As Variant()
DryzDbq = DryzRs(Rsz(A, Q))
End Function

Function DryzQ(Q) As Variant()
DryzQ = DryzDbq(CDb, Q)
End Function


Function FnyzDbq(A As Database, Q) As String()
FnyzDbq = FnyzRs(Rsz(A, Q))
End Function

Function FnyzQ(Q) As String()
FnyzQ = FnyzDbq(CDb, Q)
End Function

Private Sub ZZ_FnyzQ()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(S)
End Sub



