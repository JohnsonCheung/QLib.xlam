Attribute VB_Name = "MDao_Tbl"
Option Explicit
Const C_SkNm$ = "SecondaryKey"

Function Fny(A As Database, T, Optional NoReOpn As Boolean) As String()
Fny = Itn(ReOpnDb(A, NoReOpn).TableDefs(T).Fields)
End Function

Function ColzRs(A As Database, T, Optional F = 0) As Dao.Recordset
Set ColzRs = Rs(A, SqlSel_F_Fm(F, T))
End Function

Function ColSetzT(A As Database, T, Optional F = 0) As Aset
Set ColSetzT = ColSetzRs(ColzRs(A, T, F))
End Function

Function CntDicz(A As Database, T, F) As Dictionary
Set CntDicz = CntDiczRs(ColzRs(A, T, F))
End Function

Function Idx(A As Database, T, Nm) As Dao.Index
Set Idx = FstItrNm(A.TableDefs(T).Indexes, Nm)
End Function

Property Get HasSk(A As Database, T) As Boolean
HasSk = Not IsNothing(SkIdx(A, T))
End Property

Function HasIdx(A As Database, T, IdxNm) As Boolean
HasIdx = HasItn(A.TableDefs(T).Indexes, IdxNm)
End Function

Function FstUniqIdx(A As Database, T) As Dao.Index
Set FstUniqIdx = FstItrTrueP(A.TableDefs(T).Indexes, "Unique")
End Function

Function HasFld(A As Database, T, F) As Boolean
HasFld = HasItn(A.TableDefs(T).Fields, F)
End Function

Function HasPk(A As Database, T) As Boolean
HasPk = HasItrTrueP(A.TableDefs(T).Indexes, "Primary")
End Function

Function HasStdPk(A As Database, T) As Boolean
If Not HasPk(A, T) Then Exit Function
If Sz(PkFny(A, T)) <> 1 Then Exit Function
HasStdPk = True
End Function

Function HasIdz(A As Database, T, Id&) As Boolean
If HasPk(A, T) Then
    HasIdz = HasReczRs(RszId(A, T, Id))
End If
End Function

Function DryzTFF(A As Database, T, FF) As Variant()
DryzTFF = DryzQ(A, SqlSel_FF_Fm(FF, T))
End Function

Sub AsgColApzDrsFF(DRs As DRs, FF, ParamArray OColAp())
Dim F, J%
For Each F In FnyzFF(FF)
    OColAp(J) = ColzDrs(DRs, CStr(F))
    J = J + 1
Next
End Sub

Function RszId(A As Database, T, Id) As Dao.Recordset
Set RszId = Rs(A, SqlSel_Fm_WhId(T, Id))
End Function

Function CsvLyzDbt(A As Database, T) As String()
CsvLyzDbt = CsvLyzRs(RszT(A, T))
End Function

Function DrszT(A As Database, T) As DRs
Set DrszT = DrszRs(RszT(A, T))
End Function
Function DryzT(A As Database, T) As Variant()
DryzT = DryzRs(RszT(A, T))
End Function

Function DtzT(A As Database, T) As DT
Set DtzT = DT(T, Fny(A, T), DryzT(A, T))
End Function

Function FdStrAy(A As Database, T) As String()
Dim F, Td As Dao.TableDef
Set Td = A.TableDefs(T)
For Each F In Fny(A, T)
    PushI FdStrAy, FdStr(Td.Fields(F))
Next
End Function

Function Fds(A As Database, T) As Dao.Fields
Set Fds = A.TableDefs(T).OpenRecordset.Fields
End Function

Sub ReSeqFldzFny(A As Database, T, Fny$())
Dim F, J%, Fds As Dao.Fields
Set Fds = A.TableDefs(T).Fields
For Each F In AyReOrd(F, Fny)
    J = J + 1
    Fds(F).OrdinalPosition = J
Next
End Sub

Function SrcFbzT$(A As Database, T)
SrcFbzT = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function NColzT&(A As Database, T)
NColzT = A.TableDefs(T).Fields.Count
End Function

Function NReczDbtBexpr&(A As Database, T, Bexpr$)
NReczDbtBexpr = ValzQ(A, FmtQQ("Select Count(*) from [?]?", T, SqpWh(Bexpr)))
End Function

Function PkFny(A As Database, T) As String()
PkFny = FnyzIdx(PkIdx(A, T))
End Function

Function PkIdxNm$(A As Database, T)
PkIdxNm = ObjNm(PkIdx(A, T))
End Function

Function PkIdx(A As Database, T) As Dao.Index
Set PkIdx = FstItrTrueP(A.TableDefs(T).Indexes, "Primary")
End Function

Function RszTFF(A As Database, T, FF) As Dao.Recordset
Set RszTFF = A.OpenRecordset(SqlSel_FF_Fm(FF, T))
End Function

Function RszT(A As Database, T) As Dao.Recordset
Set RszT = A.TableDefs(T).OpenRecordset
End Function

Function FdzTF(A As Database, T, F) As Dao.Field2
Set FdzTF = A.TableDefs(T).Fields(F)
End Function

Function SqzT(A As Database, T, Optional ExlFldNm As Boolean) As Variant()
SqzT = SqzRs(RszT(A, T), ExlFldNm)
End Function

Function SrcTn$(A As Database, T)
SrcTn = A.TableDefs(T).SourceTableName
End Function

Function StruzT$(A As Database, T)
'Const CSub$ = CMod & "Stru"
'If Not Has Then FunMsgAp_Dmp_Ly CSub, "[Db] has not such [Tbl]", DbNm, T: Exit Function
Dim F$()
    F = Fny(A, T)
    If IsLnkzFx(A, T) Then
        StruzT = T & " " & JnSpc(AyQuoteSqIf(F))
        Exit Function
    End If

Dim P$
    If HasEle(F, T & "Id") Then
        P = " *Id"
        F = AyMinus(F, Array(T & "Id"))
    End If
Dim Sk, Rst
    Dim J%, X
    'Sk = SkFny
    Rst = AyMinus(F, Sk)
    If Sz(Sk) > 0 Then
        For Each X In Sk
            Sk(J) = Replace(X, T, "*")
            J = J + 1
        Next
        Sk = " " & JnSpc(AyQuoteSqIf(Sk)) & " |"
    Else
        Sk = ""
    End If
    '
    J = 0
    For Each X In Itr(Rst)
        Rst(J) = Replace(X, T, "*")
        J = J + 1
    Next
Rst = " " & JnSpc(AyQuoteSqIf(Rst))
StruzT = T & P & Sk & Rst
End Function

Function LasUpdTimz(A As Database, T) As Date
LasUpdTimz = TblPrp(A, T, "LastUpdated")
End Function

Sub InsDrsz(A As Database, T, DRs As DRs)
InsRszDry RszTFF(A, T, DRs.Fny), DRs.Dry
End Sub

Sub AddFd(A As Database, T, Fd As Dao.Fields)
A.TableDefs(T).Fields.Append Fd
End Sub

Sub AddFld(A As Database, T, F, Ty As DataTypeEnum, Optional Sz%, Optional Precious%)
If HasFld(A, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = SqlTyzDao(Ty, Sz, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
A.Execute S
End Sub

Sub RenTblz(A As Database, T, ToNm)
A.TableDefs(T).Name = ToNm
End Sub

Sub RenTblzAddPfxDbt(A As Database, T, Pfx)
RenTblz A, T, Pfx & T
End Sub

Sub BrwDbt(A As Database, T)
BrwDt DtzT(A, T)
End Sub

Property Get IsSysTbl(A As Database, T) As Boolean
IsSysTbl = (A.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbSystemObject) <> 0
End Property

Property Get IsHidTbl(A As Database, T) As Boolean
IsHidTbl = (A.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbHiddenObject) <> 0
End Property

Function LnkInf(A As Database) As String()
Dim T
For Each T In Tni(A)
    PushI LnkInf, LnkInfzT(A, T)
Next
End Function

Function LnkInfzT$(A As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(A, T): LnkInfzT = FmtQQ("LnkFx(?).LnkWs(?).Tbl(?).Db(?)", DtaSrc(A, T), SrcTn(A, T), T, DbNm(A))
Case IsLnkzFb(A, T): LnkInfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", DtaSrc(A, T), SrcTn(A, T), T, DbNm(A))
End Select
End Function

Function Acs() As Access.Application
Static A As Access.Application
If IsNothing(A) Then Set A = New Access.Application: A.Visible = True
Set Acs = A
End Function

Function CrtTblzDupKey$(A As Database, Into, FmTbl, KK$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SySsl(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into, FmTbl, Tmp, Jn)
DrpT A, Tmp
End Function

Sub InsTblzDry(A As Database, T, Dry())
InsRszDry RszT(A, T), Dry
End Sub

Sub CrtTblzJnFld(A As Database, T, KK, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
RunQ A, SqlSel_FF_Into_Fm_WhFalse(KK, Tar, T)
AddFld A, T, LisFld, dbMemo
InsTblzDry A, T, DryzJnFldKK(DryzT(A, T), KK, FldIx(A, T, JnFld))
End Sub

Function FldIx%(A As Database, T, Fld)
Dim F As Dao.Field, O%
For Each F In A.TableDefs(T).Fields
    If F.Name = Fld Then
        FldIx = O
        Exit Function
    End If
    O = O + 1
Next
FldIx = -1
End Function
Sub CrtPk(A As Database, T)
A.Execute SqlCrtPkzT(T)
End Sub

Function JnQSqCommaSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnQSqCommaSpcAp = JnQSqCommaSpc(Av)
End Function
Function CommaSpcSqAv$(Av())

End Function
Function JnCommaSpcFF$(FF)
JnCommaSpcFF = JnQSqCommaSpc(FnyzFF(FF))
End Function

Sub CrtSk(A As Database, T, SkFF)
A.Execute SqlCrtSk(T, SkFF)
End Sub

Sub DrpFld(A As Database, T, FF)
Dim F
For Each F In ItrTT(FF)
    A.Execute SqlDrpCol_T_F(T, F)
Next
End Sub

Sub RenFld(A As Database, T, F, ToFld)
A.TableDefs(T).Fields(F).Name = ToFld
End Sub

Sub UpdValIdFldz(A As Database, T, ValFld, ValIdFld)
Dim D As New Dictionary, J&, Rs As Dao.Recordset, V
Set Rs = Rs(SqlSel_X_Fm(JnQSqCommaSpcAp(ValFld, ValIdFld), T))
With Rs
    While Not .EOF
        .Edit
        V = .Fields(0).Value
        If D.Exists(V) Then
            .Fields(1).Value = D(V)
        Else
            .Fields(1).Value = J
            D.Add V, J
            J = J + 1
        End If
        .Update
        .MoveNext
    Wend
End With
End Sub

Function FdzFld(A As Database, T, Fld) As Dao.Field2
Set FdzFld = A.TableDefs(T).Fields(Fld)
End Function

Function FdStrzTF$(A As Database, T, F)
FdStrzTF = FdStr(FdzTF(A, T, F))
End Function

Function IntAyzDbtf(A As Database, T, F) As Integer()
Q = FmtQQ("Select [?] from [?]", F, T)
IntAyzDbtf = IntAyzDbq(A, Q)
End Function

Function IsPkFld(F) As Boolean
'IsPkFld = HasEle(PkFny, F)
End Function

Function NxtId&()
'Dim S$: S = FmtQQ("select Max(?Id) from ?", T, T)
'NxtId = ValzQ(S) + 1
End Function

Function SyFld(F) As String()
'SyFld = IntozRs(SyFld, ColRs(F))
End Function

Function DaoTyFld(F) As Dao.DataTypeEnum
'DaoTyFld = A.TableDefs(T).Fields(F).Type
End Function

Function ShtTyzDaoFld$(F)
ShtTyzDaoFld = ShtTyzDao(DaoTyFld(F))
End Function

Property Get LnkTblCnStr$()
On Error Resume Next
'LnkTblCnStr = A.TableDefs(T).Connect
End Property
Sub AddExprFld(F, Expr$, Ty As Dao.DataTypeEnum)
'A.TableDefs(T).Fields.Append Fd(F, Ty, Expr:=Expr)
End Sub

Function ValRecIdFld(RecId&, Fld)  ' K is Pk value
'ValRecIdFld = ValzQ(SqlSel_FF_Fm(T, Fld, BexprRecId(T, RecId)))
End Function
Sub CrtTblzEmpClone(TblToCrt)
'Run SqlSel_Into_Fm_WhFalse(TblToCrt, T)
End Sub

Sub KillTmpDb()

End Sub
Private Sub Z_CrtDupKeyTbl()
Dim D As Database: Set D = TmpDb
DrpTT D, "#A #B"
'T = "AA"
CrtTblzDupKey D, "#A", "#B", "Sku BchNo"
DrpDbIfTmp D
End Sub

Private Sub Z_PkFny()
ZZ:
    Dim A As Database
    Set A = Db(SampFbzDutyDta)
    Dim Dr(), Dry(), T
    For Each T In Tny(A)
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFny(A, T)
        PushI Dry, Dr
    Next
    BrwDry Dry
    Exit Sub
End Sub

Private Sub ZZ()
Dim A As Database
Dim B
Dim C As Dao.Fields
Dim D As DataTypeEnum
Dim E%
Dim F$
Dim G As Boolean
Dim H()
Dim I$()
Dim J&
Dim L As Dictionary
Dim M As Dao.Index
Dim O As Dao.Database
Dim XX
End Sub

Property Get NReczT&(A As Database, T)
NReczT = ValzQ(A, SqlSelCnt_T(T))
End Property

Function ValzQQ(A As Database, QQSql, ParamArray Ap())
Dim Av(): Av = Ap
ValzQQ = ValzQ(A, FmtQQAv(QQSql, Av))
End Function

Property Get LoFmtrVblPrp$(A As Database, T)
LoFmtrVblPrp = TblPrp(A, T, "LoFmtrVbl")
End Property

Property Let LoFmtrVblPrp(A As Database, T, LoFmtrVbl$)
TblPrp(A, T, "LoFmtrVbl") = LoFmtrVbl
End Property

Function IsLnk(A As Database, T) As Boolean
IsLnk = IsLnkzFb(A, T) Or IsLnkzFx(A, T)
End Function

Function CnStrzT$(A As Database, T)
On Error Resume Next
CnStrzT = A.TableDefs(T).Connect
End Function
Property Get IsLnkzFb(A As Database, T) As Boolean
IsLnkzFb = HasPfx(CnStrzT(A, T), ";Database=")
End Property

Function IsLnkzFx(A As Database, T) As Boolean
IsLnkzFx = HasPfx(CnStrzT(A, T), "Excel")
End Function
Private Sub Z_AddExprFld()
'DrpTT "Tmp"
Dim A As Dao.TableDef
'Set A = AddTd(CDb, TmpTd)
'AddDbtfExpr CDb, "Tmp", "F2", "[F1]+"" hello!"""
'DrpTT "Tmp"
End Sub


