Attribute VB_Name = "QDao_Db_Get"
Option Explicit
Private Const CMod$ = "MDao_Db_Get."
Private Const Asm$ = "QDao"

Function LngAyzQ(A As Database, Q) As Long()
LngAyzQ = LngAyzRs(Rs(A, Q))
End Function

Function SyzQ(A As Database, Q) As String()
SyzQ = SyzRs(Rs(A, Q))
End Function

Private Sub ZZ_Rs()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsvLyzRs(Rs(TmpDb, S))
End Sub



