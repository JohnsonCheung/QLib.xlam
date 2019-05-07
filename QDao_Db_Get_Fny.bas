Attribute VB_Name = "QDao_Db_Get_Fny"
Option Explicit
Private Const CMod$ = "MDao_Db_Get_Fny."
Private Const Asm$ = "QDao"

Function FnyzQ(A As Database, Q$) As String()
FnyzQ = FnyzRs(Rs(A, Q))
End Function

Private Sub Z_FnyzQ()
Dim Db As Database
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(Db, S)
End Sub



