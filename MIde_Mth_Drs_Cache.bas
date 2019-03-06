Attribute VB_Name = "MIde_Mth_Drs_Cache"
Option Explicit

Function CacheDtevPjf(Pjf) As Date
CacheDtevPjf = ValzQ(MthDb, FmtQQ("Select PjDte from Mth where Pjf='?'", Pjf))
End Function

Function MthDrszPjfzFmCache(Pjf, Optional WhStr$) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where Pjf='?'", Pjf)
Set MthDrszPjfzFmCache = DrszFbq(MthDb, Sql)
End Function
Sub EnsTblMthCachezPjf(Pjf)
Dim D1 As Date
Dim D2 As Date
    D1 = PjDtePjf(Pjf)
    D2 = CacheDtevPjf(Pjf)
Select Case True
Case D1 = 0:  Stop
Case D2 = 0:
Case D1 = D2: Exit Sub
Case D2 > D1: Stop
End Select
Stop '
'DrsIup_Dbt MthDRszPjf(Pjf), MthDb, "MthCache", FmtQQ("Pjf='?'", Pjf)
End Sub

