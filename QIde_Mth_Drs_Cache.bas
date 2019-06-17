Attribute VB_Name = "QIde_Mth_Drs_Cache"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_DCache."
Private Const Asm$ = "QIde"
Public Const DoczTof$ = "DashNm: tbl-of.  IsCml. After Tof, it is a table-name."

Function CacheDtezPjf(Pjf) As Date
CacheDtezPjf = ValzQ(MthDbP, FmtQQ("Select PjDte from Mth where Pjf='?'", Pjf))
End Function

Function DMthzPjfzFmCache(Pjf) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where Pjf='?'", Pjf)
DMthzPjfzFmCache = DrszFbq(MthDbP, Sql)
End Function

Function SkFnyWiSqlQPfx(A As Database, T) As String()
Dim F
For Each F In Itr(SkFny(A, T))
    PushI SkFnyWiSqlQPfx, SqlQuoteChrzT(DaoTyzTF(A, T, F)) & F
Next
End Function
Sub IupDbt(A As Database, T, Drs As Drs)
Dim Dry(): Dry = Drs.Dry
If Si(Dry) = 0 Then Exit Sub
'ThwIf_DrsGoodToIupDbt CSub, Drs, A, T
Dim R As Dao.Recordset, Q$, Sql$, Dr
'Sql = SqlSel_T_Wh(T, BexpzFnyzSqlQPfxSy(SkFny(A, T), SkSqlQPfxSy(A, T)))
For Each Dr In Dry
    Q = FmtQQAv(Sql, CvAv(Dr))
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

