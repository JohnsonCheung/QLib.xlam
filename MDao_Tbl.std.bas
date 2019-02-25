Attribute VB_Name = "MDao_Tbl"
Option Explicit
Const C_SkNm$ = "SecondaryKey"
Function Fny(T, Optional NoReOpn As Boolean) As String()
Fny = Fnyz(CDb, T, NoReOpn)
End Function

Function Fnyz(A As Database, T, Optional NoReOpn As Boolean) As String()
Fnyz = Itn(Dbz(A, NoReOpn).TableDefs(T).Fields)
End Function

Function ColRsz(A As Database, T, Optional F = 0) As DAO.Recordset
Set ColRsz = Rsz(A, SqlSel_F_Fm(F, T))
End Function

Function ColSetz(A As Database, T, Optional F = 0) As Aset
Set ColSetz = ColSetzRs(ColRsz(A, T, F))
End Function

Function CntDicz(A As Database, T, F) As Dictionary
Set CntDicz = CntDiczRs(ColRsz(A, T, F))
End Function

Function Idxz(A As Database, T, Nm) As DAO.Index
Set Idxz = FstItrNm(A.TableDefs(T).Indexes, Nm)
End Function

Property Get HasSk() As Boolean
'HasSk = Not IsNothing(SkIdx)
End Property

Function HasIdxz(A As Database, T, IdxNm) As Boolean
HasIdxz = HasItn(A.TableDefs(T).Indexes, IdxNm)
End Function

Function FstUniqIdxz(A As Database, T) As DAO.Index
Set FstUniqIdxz = FstItrTrueP(A.TableDefs(T).Indexes, "Unique")
End Function

Function HasFldz(A As Database, T, F) As Boolean
HasFldz = HasItn(A.TableDefs(T).Fields, F)
End Function

Function HasPkz(A As Database, T) As Boolean
HasPkz = HasItrTrueP(A.TableDefs(T).Indexes, "Primary")
End Function

Function HasStdPkz(A As Database, T) As Boolean
If Not HasPkz(A, T) Then Exit Function
If Sz(PkFnyz(A, T)) <> 1 Then Exit Function
HasStdPkz = True
End Function

Function HasIdz(A As Database, T, Id&) As Boolean
If HasPkz(A, T) Then
    HasIdz = HasReczRs(RszId(A, T, Id))
End If
End Function

Function DryzDbtFF(A As Database, T, FF) As Variant()
DryzDbtFF = DryzDbq(A, SqlSel_FF_Fm(FF, T))
End Function

Sub AsgColApzDrsFF(Drs As Drs, FF, ParamArray OColAp())
Dim F, J%
For Each F In FnyzFF(FF)
    OColAp(J) = ColzDrs(Drs, CStr(F))
    J = J + 1
Next
End Sub

Function RszId(A As Database, T, Id) As DAO.Recordset
Set RszId = Rsz(A, SqlSel_Fm_WhId(T, Id))
End Function

Function PkIdxz(A As Database, T) As DAO.Index
Set PkIdxz = FstItrTrueP(A.TableDefs(T).Indexes, "Primary")
End Function

Function CsvLyzDbt(A As Database, T) As String()
CsvLyzDbt = CsvLyzRs(RszDbt(A, T))
End Function

Function DrszDbt(A As Database, T) As Drs
Set DrszDbt = DrszRs(RszDbt(A, T))
End Function
Function Dryz(T) As Variant()
Dryz = DryzT(CDb, T)
End Function
Function DryzT(A As Database, T) As Variant()
DryzT = DryzRs(RszDbt(A, T))
End Function

Function DtzDbt(A As Database, T) As Dt
Set DtzDbt = Dt(T, Fnyz(A, T), DryzT(A, T))
End Function

Function FdStrAy(A As Database, T) As String()
Dim F, Td As DAO.TableDef
Set Td = A.TableDefs(T)
For Each F In Fnyz(A, T)
    PushI FdStrAy, FdStr(Td.Fields(F))
Next
End Function

Function Fds(A As Database, T) As DAO.Fields
Set Fds = A.TableDefs(T).OpenRecordset.Fields
End Function

Sub ReSeqFldzFny(A As Database, T, Fny$())
Dim F, J%, Fds As DAO.Fields
Set Fds = A.TableDefs(T).Fields
For Each F In AyReOrd(F, Fny)
    J = J + 1
    Fds(F).OrdinalPosition = J
Next
End Sub

Function SrcFbzDbt$(A As Database, T)
SrcFbzDbt = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function NColzDbt&(A As Database, T)
NColzDbt = A.TableDefs(T).Fields.Count
End Function

Function NReczDbtBexpr&(A As Database, T, Bexpr$)
NReczDbtBexpr = ValzQ(FmtQQ("Select Count(*) from [?]?", T, SqpWh(Bexpr)))
End Function

Function PkFnyz(A As Database, T) As String()
PkFnyz = FnyzIdx(PkIdxz(A, T))
End Function

Function PkIdxNm$(A As Database, T)
PkIdxNm = ObjNm(PkIdxz(A, T))
End Function

Function RszDbtFF(A As Database, T, FF) As DAO.Recordset
Set RszDbtFF = A.OpenRecordset(SqlSel_FF_Fm(FF, T))
End Function

Function RszDbt(A As Database, T) As DAO.Recordset
Set RszDbt = A.OpenRecordset(T)
End Function

Function Fd(A As Database, T, F) As DAO.Field2
Set Fd = A.TableDefs(T).Fields(F)
End Function

Function SqzDbt(A As Database, T, Optional ExlFldNm As Boolean) As Variant()
SqzDbt = SqzRs(RszDbt(A, T), ExlFldNm)
End Function

Function SrcTz$(A As Database, T)
SrcTz = A.TableDefs(T).SourceTableName
End Function

Function StruzT$(A As Database, T)
'Const CSub$ = CMod & "Stru"
'If Not Has Then FunMsgAp_Dmp_Ly CSub, "[Db] has not such [Tbl]", DbNm, T: Exit Function
Dim F$()
    F = Fnyz(A, T)
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
LasUpdTimz = TblPrpz(A, T, "LastUpdated")
End Function

Sub InsDrsz(A As Database, T, Drs As Drs)
InsRszDry RszDbtFF(A, T, Drs.Fny), Drs.Dry
End Sub

Sub AddFd(A As Database, T, Fd As DAO.Fields)
A.TableDefs(T).Fields.Append Fd
End Sub

Sub AddFld(A As Database, T, F, Ty As DataTypeEnum, Optional Sz%, Optional Precious%)
If HasFldz(A, T, F) Then Exit Sub
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
BrwDt DtzDbt(A, T)
End Sub

Property Get IsSysTbl(A As Database, T) As Boolean
IsSysTbl = (A.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject) <> 0
End Property

Property Get IsHidTbl(A As Database, T) As Boolean
IsHidTbl = (A.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbHiddenObject) <> 0
End Property
Property Get LnkInf() As String()
LnkInf = LnkInfz(CDb)
End Property

Function DtaSrcz$(A As Database, T)

End Function
Function LnkInfz(A As Database) As String()

End Function

Function LnkInfzT$(A As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(A, T): LnkInfzT = FmtQQ("LnkFx(?).LnkWs(?).Tbl(?).Db(?)", DtaSrcz(A, T), SrcTz(A, T), T, DbNm(A))
Case IsLnkzFb(A, T): LnkInfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", DtaSrcz(A, T), SrcTz(A, T), T, DbNm(A))
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
Drpz A, Tmp
End Function

Sub InsTblzDry(A As Database, T, Dry())
InsRszDry RszDbt(A, T), Dry
End Sub

Sub CrtTblzJnFld(A As Database, T, KK, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
RunQz A, SqlSel_FF_Into_Fm_WhFalse(KK, Tar, T)
AddFld A, T, LisFld, dbMemo
InsTblzDry A, T, DryzJnFldKK(DryzT(A, T), KK, FldIx(A, T, JnFld))
End Sub

Function FldIx%(A As Database, T, Fld)
Dim F As DAO.Field, O%
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
A.Execute SqlCrtPk_T(T)
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

Function SqlCrtSk$(A As Database, T, SkFF)
SqlCrtSk = FmtQQ("Create Unique Index SecondaryKey on ? (?)", T, JnCommaSpcFF(SkFF))
End Function
Sub CrtSk(A As Database, T, SkFF)
A.Execute SqlCrtSk(A, T, SkFF)
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
Dim D As New Dictionary, J&, Rs As DAO.Recordset, V
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

Function FdzFld(A As Database, T, Fld) As DAO.Field2
Set FdzFld = A.TableDefs(T).Fields(Fld)
End Function

Function FdStrzDbtf$(A As Database, T, F)
FdStrzDbtf = FdStr(Fd(A, T, F))
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

Function DaoTyFld(F) As DAO.DataTypeEnum
'DaoTyFld = A.TableDefs(T).Fields(F).Type
End Function

Function ShtTyzDaoFld$(F)
ShtTyzDaoFld = ShtTyzDao(DaoTyFld(F))
End Function

Property Get LnkTblCnStr$()
On Error Resume Next
'LnkTblCnStr = A.TableDefs(T).Connect
End Property
Sub AddExprFld(F, Expr$, Ty As DAO.DataTypeEnum)
'A.TableDefs(T).Fields.Append NewFd(F, Ty, Expr:=Expr)
End Sub

Function ValRecIdFld(RecId&, Fld)  ' K is Pk value
'ValRecIdFld = ValzQ(SqlSel_FF_Fm(T, Fld, BexprRecId(T, RecId)))
End Function
Sub CrtTblzEmpClone(TblToCrt)
'Run SqlSel_Into_Fm_WhFalse(TblToCrt, T)
End Sub
Sub CrtTblzDupKeyDb(A As Database, T, KK)
CrtTblzEmpClone T
'Dbx.Tbl(T).InsDry DryDupKeyKK(DryzT(A, T), KK)
End Sub

Sub KillTmpDb()

End Sub
Private Sub Z_CrtDupKeyTbl()
Drp "#A #B"
Dim D As Database: Set D = TmpDb
'T = "AA"
CrtTblzDupKeyDb D, "#A", "Sku BchNo"
DrpDbIfTmpz D
End Sub

Private Sub Z_PkFny()
ZZ:
    Dim A As Database
    Set A = Db(SampFbzDutyDta)
    Dim Dr(), Dry(), T
    For Each T In Tnyz(A)
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFnyz(A, T)
        PushI Dry, Dr
    Next
    BrwDry Dry
    Exit Sub
End Sub

Private Sub ZZ()
Dim A As Database
Dim B
Dim C As DAO.Fields
Dim D As DataTypeEnum
Dim E%
Dim F$
Dim G As Boolean
Dim H()
Dim I$()
Dim J&
Dim L As Dictionary
Dim M As DAO.Index
Dim O As DAO.Database
Dim XX
End Sub

Private Sub Z()
End Sub
Function NRec&(T)
NRec = NRecz(CDb, T)
End Function
Property Get NRecz&(A As Database, T)
NRecz = ValzDbq(A, SqlSelCnt_T(T))
End Property

Function ValzDbQQ(A As Database, QQSql, ParamArray Ap())
Dim Av(): Av = Ap
ValzDbQQ = ValzDbq(A, FmtQQAv(QQSql, Av))
End Function

Property Get LoFmtrVblPrpz$(A As Database, T)
LoFmtrVblPrpz = TblPrpz(A, T, "LoFmtrVbl")
End Property

Property Let LoFmtrVblPrpz(A As Database, T, LoFmtrVbl$)
TblPrpz(A, T, "LoFmtrVbl") = LoFmtrVbl
End Property

Property Get LoFmtrVblPrp$(T)
LoFmtrVblPrp = LoFmtrVblPrpz(CDb, T)
End Property

Property Let LoFmtrVblPrp(T, LoFmtrVbl$)
TblPrp(T, "LoFmtrVbl") = LoFmtrVbl
End Property

Function IsLnk(T) As Boolean
IsLnk = IsLnkz(CDb, T)
End Function

Function IsLnkz(A As Database, T) As Boolean
IsLnkz = IsLnkzFb(A, T) Or IsLnkzFx(A, T)
End Function

Function CnStr$(T)
CnStr = CnStrz(CDb, T)
End Function
Function CnStrz$(A As Database, T)
On Error Resume Next
CnStrz = A.TableDefs(T).Connect
End Function
Property Get IsLnkzFb(A As Database, T) As Boolean
IsLnkzFb = HasPfx(CnStrz(A, T), ";Database=")
End Property

Function IsLnkzFx(A As Database, T) As Boolean
IsLnkzFx = HasPfx(CnStrz(A, T), "Excel")
End Function
Private Sub Z_AddExprFld()
'DrpTT "Tmp"
Dim A As DAO.TableDef
'Set A = AddTd(CDb, TmpTd)
'AddDbtfExpr CDb, "Tmp", "F2", "[F1]+"" hello!"""
'DrpTT "Tmp"
End Sub


