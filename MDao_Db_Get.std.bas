Attribute VB_Name = "MDao_Db_Get"
Option Explicit

Function LngAyzDbq(A As Database, Q) As Long()
LngAyzDbq = LngAyzRs(Rsz(A, Q))
End Function

Function LngAyzSql(Q) As Long()
LngAyzSql = LngAyzDbq(CDb, Q)
End Function

Function SyzDbq(A As Database, Q) As String()
SyzDbq = SyzRs(Rsz(A, Q))
End Function

Function SyzQ(Q) As String()
SyzQ = SyzDbq(CDb, Q)
End Function

Private Sub ZZ_RszSql()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsvLyzRs(Rs(S))
End Sub

Private Sub Z_SyzQ()
DmpAy SyzQ("Select Distinct UOR from [>Imp]")
End Sub


