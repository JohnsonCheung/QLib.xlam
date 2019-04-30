Attribute VB_Name = "MDao_Db_Get_Fny"
Option Explicit

Function FnyzQ(A As Database, Q$) As String()
FnyzQ = FnyzRs(Rs(A, Q))
End Function

Private Sub Z_FnyzQ()
Dim Db As Database
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(Db, S)
End Sub



