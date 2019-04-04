Attribute VB_Name = "MIde_Mth_Drs_Cache"
Option Explicit
Public Const DoczTof$ = "DashNm: tbl-of.  IsCml. After Tof, it is a table-name."

Function CacheDtevPjf(Pjf) As Date
CacheDtevPjf = ValOfQ(MthDbOfPj, FmtQQ("Select PjDte from Mth where Pjf='?'", Pjf))
End Function

Function MthDrszPjfzFmCache(Pjf, Optional WhStr$) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where Pjf='?'", Pjf)
Set MthDrszPjfzFmCache = DrszFbq(MthDbOfPj, Sql)
End Function
Sub EnsTofMthCachezPjf(Pjf)
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
IupDbt MthDb, "MthCache", MthDrszPjf(Pjf, FmtQQ("Pjf='?'", Pjf))
End Sub
Sub ThwIfDrsGoodToIupDbt(Drs As Drs, Db As Database, T)

End Sub

Function BexprzFnyWiSqlQuMk(FNyWiSqlQuMk$())
End Function
Function SqlQuMk$(A As Dao.DataTypeEnum)
Select Case A
Case _
    Dao.DataTypeEnum.dbBigInt, _
    Dao.DataTypeEnum.dbByte, _
    Dao.DataTypeEnum.dbCurrency, _
    Dao.DataTypeEnum.dbDecimal, _
    Dao.DataTypeEnum.dbDouble, _
    Dao.DataTypeEnum.dbFloat, _
    Dao.DataTypeEnum.dbInteger, _
    Dao.DataTypeEnum.dbLong, _
    Dao.DataTypeEnum.dbNumeric, _
    Dao.DataTypeEnum.dbSingle: Exit Function
Case _
    Dao.DataTypeEnum.dbChar, _
    Dao.DataTypeEnum.dbMemo, _
    Dao.DataTypeEnum.dbText: SqlQuMk = "'"
Case _
    Dao.DataTypeEnum.dbDate: SqlQuMk = "#"
Case Else
    Thw CSub, "Invalid DaoTy", "DaoTy", A
End Select
End Function
Function SkFnyWiSqlQuMkPfx(A As Database, T) As String()
Dim F
For Each F In Itr(SkFny(A, T))
    PushI SkFnyWiSqlQuMkPfx, SqlQuMk(DaoTyzTF(A, T, F)) & F
Next
End Function
Sub IupDbt(A As Database, T, Drs As Drs)
Dim Dry(): Dry = Drs.Dry
If Si(Dry) = 0 Then Exit Sub
ThwIfDrsGoodToIupDbt CSub, Drs, A, T
Dim R As Dao.Recordset, Q$, Sql$
Sql = SqlSel_T_Wh(T, BexprzFnyWiSqlQuMkPfx(SkFnyWiSqlQuMkPfx(A, T)))
For Each Dr In Dry
    Q = FmtQQAv(Sql, Dr)
    Set R = Rs(A, Q)
    If HasRec(R) Then
        UpdRs R, Dr
    Else
        InsRs R, Dr
    End If
Next
End Sub
Sub InsDbt(A As Database, T, Dry())

End Sub
Sub UpdDbt(A As Database, T, Dry())

End Sub

