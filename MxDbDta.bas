Attribute VB_Name = "MxDbDta"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDbDta."

Function LngAyzQ(D As Database, Q) As Long()
LngAyzQ = LngAyzRs(Rs(D, Q))
End Function

Function SyzQ(D As Database, Q) As String()
SyzQ = SyzRs(Rs(D, Q))
End Function

Private Sub Z_Rs()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsvLyzRs(Rs(TmpDb, S))
End Sub


Sub BrwQ(D As Database, Q)
BrwDrs DrszQ(D, Q)
End Sub

Function IntAyzQ(D As Database, Q) As Integer()
End Function

Function SyzTF(D As Database, T, F$) As String()
SyzTF = SyzRs(RszTF(D, T, F))
End Function

Function IntozTF(Into, D As Database, T, F$)
IntozTF = IntozRs(Into, RszTF(D, T, F))
End Function



Function FnyzQ(D As Database, Q) As String()
FnyzQ = FnyzRs(Rs(D, Q))
End Function

Private Sub Z_FnyzQ()
Dim Db As Database
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(Db, S)
End Sub






Function DrzQ(D As Database, Q) As Variant()
DrzQ = DrzRs(Rs(D, Q))
End Function

Function DyoQ(D As Database, Q) As Variant()
DyoQ = DyoRs(Rs(D, Q))
End Function