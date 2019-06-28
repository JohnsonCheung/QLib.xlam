Attribute VB_Name = "QDao_Tbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tbl."
Private Const Asm$ = "QDao"
Const C_SkNm$ = "SecondaryKey"

Function FnyzT(D As Database, T) As String()
FnyzT = Fny(D, T)
End Function

Function Fny(D As Database, T) As String()
Fny = Itn(DbzReOpn(D).TableDefs(T).Fields)
End Function

Function ColzRs(D As Database, T, F$) As Dao.Recordset
Set ColzRs = Rs(D, SqlSel_F_T(F, T))
End Function

Function ColSetzT(D As Database, T, F$) As Aset
Set ColSetzT = ColSetzRs(ColzRs(D, T, F))
End Function

Function DiKqCntzTF(D As Database, T, F$) As Dictionary
Set DiKqCntzTF = DiKqCntzRs(ColzRs(D, T, F$))
End Function

Function IdxzTd(A As Dao.TableDef, IdxNm$) As Dao.Index
Set IdxzTd = FstzItrNm(A.Indexes, IdxNm$)
End Function

Function Idx(D As Database, T, IdxNm$) As Dao.Index
Set Idx = IdxzTd(D.TableDefs(T), IdxNm)
End Function

Function HasSk(D As Database, T) As Boolean
HasSk = Not IsNothing(SkIdx(D, T))
End Function

Function HasIdx(D As Database, T, IdxNm$) As Boolean
HasIdx = HasItn(D.TableDefs(T).Indexes, IdxNm)
End Function

Function FstUniqIdx(D As Database, T) As Dao.Index
Set FstUniqIdx = FstzItrT(D.TableDefs(T).Indexes, "Unique")
End Function

Function HasFld(D As Database, T, F$) As Boolean
HasFld = HasItn(D.TableDefs(T).Fields, F)
End Function

Function HasPk(D As Database, T) As Boolean
HasPk = HasPkzTd(D.TableDefs(T))
End Function

Function HasPkzTd(A As Dao.TableDef) As Boolean
HasPkzTd = HasItn(A.Indexes, C_PkNm)
End Function

Function HasStdSkzTd(A As Dao.TableDef) As Boolean
If Not HasItn(A.Indexes, C_SkNm) Then Exit Function
HasStdSkzTd = A.Indexes(C_SkNm).Unique = True
End Function

Function HasStdPkzTd(A As Dao.TableDef) As Boolean
If Not HasPkzTd(A) Then Exit Function
Dim Pk$(): Pk = PkFnyzTd(A): If Si(Pk) <> 1 Then Exit Function
Dim P$: P = A.Name & "Id"
If Pk(0) <> P Then Exit Function
HasStdPkzTd = A.Fields(0).Name <> P
End Function

Function HasStdPk(D As Database, T) As Boolean
HasStdPk = HasStdPkzTd(D.TableDefs(T))
End Function

Function HasId(D As Database, T, Id&) As Boolean
If HasPk(D, T) Then
    HasId = HasRec(RszId(D, T, Id))
End If
End Function

Function DyoTFF(D As Database, T, FF$) As Variant()
DyoTFF = DyoQ(D, SqlSel_FF_T(FF, T))
End Function

Sub AsgColApzDrsFF(D As Drs, FF$, ParamArray OColAp())
Dim F, J%
For Each F In TermAy(FF)
    OColAp(J) = ColzDrs(D, CStr(F))
    J = J + 1
Next
End Sub

Function RszId(D As Database, T, Id&) As Dao.Recordset
Set RszId = Rs(D, SqlSel_T_WhId(T, Id))
End Function

Function CsvLyzDbt(D As Database, T) As String()
CsvLyzDbt = CsvLyzRs(RszT(D, T))
End Function

Function DrszT(D As Database, T) As Drs
DrszT = DrszRs(RszT(D, T))
End Function
Function DyoT(D As Database, T) As Variant()
DyoT = DyoRs(RszT(D, T))
End Function

Function DtzT(D As Database, T) As DT
DtzT = DT(T, Fny(D, T), DyoT(D, T))
End Function

Function FdStrAy(D As Database, T) As String()
Dim F, Td As Dao.TableDef
Set Td = D.TableDefs(T)
For Each F In Fny(D, T)
    PushI FdStrAy, FdStr(Td.Fields(F))
Next
End Function

Function Fds(D As Database, T) As Dao.Fields
Set Fds = D.TableDefs(T).OpenRecordset.Fields
End Function

Sub ReSeqFldzFny(D As Database, T, Fny$())
Dim F, J%, Fds As Dao.Fields
Set Fds = D.TableDefs(T).Fields
For Each F In ReOrdAy(F, Fny)
    J = J + 1
    Fds(F).OrdinalPosition = J
Next
End Sub

Function SrcFbzT$(D As Database, T)
SrcFbzT = Bet(D.TableDefs(T).Connect, "Database=", ";")
End Function

Function NColzT&(D As Database, T)
NColzT = D.TableDefs(T).Fields.Count
End Function

Function NReczTBexp&(D As Database, T, Bexp$)
NReczTBexp = VzQ(D, SqlSelCnt_T_OB(T, Bexp))
End Function

Function PkFnyzTd(A As Dao.TableDef) As String()
PkFnyzTd = FnyzIdx(PkizTd(A))
End Function

Function PkFny(D As Database, T) As String()
PkFny = FnyzIdx(PkIdx(D, T))
End Function

Function PkIdxNm$(D As Database, T)
PkIdxNm = ObjNm(PkIdx(D, T))
End Function

Function PkizTd(A As Dao.TableDef) As Dao.Index
Set PkizTd = FstzItn(A.Indexes, C_PkNm)
End Function

Function PkIdx(D As Database, T) As Dao.Index
Set PkIdx = PkizTd(D.TableDefs(T))
End Function

Function RszTFny(D As Database, T, Fny$()) As Dao.Recordset
Set RszTFny = D.OpenRecordset(SqlSel_Fny_T(Fny, T))
End Function

Function RszTFF(D As Database, T, FF$) As Dao.Recordset
Set RszTFF = RszTFny(D, T, Ny(FF))
End Function

Function RszTF(D As Database, T, F$) As Dao.Recordset
Set RszTF = D.OpenRecordset(SqlSel_F_T(F, T))
End Function

Function RszT(D As Database, T) As Dao.Recordset
Set RszT = Rs(D, SqlSel_T(T))
End Function

Function FdzTF(D As Database, T, F$) As Dao.Field2
Set FdzTF = D.TableDefs(T).Fields(F)
End Function

Function SqzT(D As Database, T, Optional ExlFldNm As Boolean) As Variant()
SqzT = SqzRs(RszT(D, T), ExlFldNm)
End Function

Function SrcTn$(D As Database, T)
SrcTn = D.TableDefs(T).SourceTableName
End Function

Function StruzT$(D As Database, T)
'Const CSub$ = CMod & "Stru"
'If Not Has Then FunMsgAp_Dmp_Ly CSub, "[Db] has not such [Tbl]", Dbn, T: Exit Function
Dim F$()
    F = Fny(D, T)
    If IsLnkzFx(D, T) Then
        StruzT = T & " " & JnSpc(SyzQteSqIf(F))
        Exit Function
    End If

Dim P$
    If HasEle(F, T & "Id") Then
        P = " *Id"
        F = MinusAy(F, Array(T & "Id"))
    End If
Dim Sk$()
    Sk = SkFny(D, T)

Dim R$
    Dim I
    Dim Rst$()
    Rst = RplStarzAy(MinusAy(F, Sk), T)
    R = " " & JnSpc(SyzQteSqIf(Rst))

Dim S$
    S = JnSpc(SyzQteSqIf(RplStarzAy(Sk, T)))
    If S <> "" Then S = " " & S & " |"

StruzT = T & P & S & R
End Function

Function LasUpdTim(D As Database, T) As Date
LasUpdTim = TblPrp(D, T, "LastUpdated")
End Function

Sub InsTblzDrs(D As Database, T, B As Drs)
InsRszDy RszTFny(D, T, B.Fny), B.Dy
End Sub

Sub AddFd(D As Database, T, Fd As Dao.Fields)
D.TableDefs(T).Fields.Append Fd
End Sub

Sub AddFld(D As Database, T, F$, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
If HasFld(D, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = SqlTyzDao(Ty, Si, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
D.Execute S
End Sub

Sub RenTblzAddPfx(D As Database, T, Pfx$)
RenTbl D, T, Pfx & T
End Sub

Sub BrwTblzByDt(D As Database, T)
BrwDt DtzT(D, T)
End Sub

Function IsSysTbl(D As Database, T) As Boolean
IsSysTbl = (D.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbSystemObject) <> 0
End Function

Function IsHidTbl(D As Database, T) As Boolean
IsHidTbl = (D.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbHiddenObject) <> 0
End Function

Function Lnkinf(D As Database) As String()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushI Lnkinf, LnkinfzT(D, T)
Next
End Function

Function LnkinfzT$(D As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(D, T): LnkinfzT = FmtQQ("LnkFx(?).LnkWs(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
Case IsLnkzFb(D, T): LnkinfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
End Select
End Function

Sub CrtTzDup(D As Database, T, FmTbl, KK$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SyzSS(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
Dim Into$
D.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
D.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into, FmTbl, Tmp, Jn)
DrpT D, Tmp
End Sub

Private Sub Z_CrtTzDrs()
Dim D As Database
GoSub Z
Exit Sub
Z:
    Set D = TmpDb
    DrpTmp D
    CrtTzDrs D, "#D", SampDrs
    BrwDb D
    Return
End Sub

Sub CrtTzDrs(D As Database, T, Drs As Drs)
CrtTblOfEmpzDrs D, T, Drs
InsTblzDy D, T, Drs.Dy
End Sub

Sub CrtTzDrszAllStr(D As Database, T, Drs As Drs)
CrtTzDrszEmpzAllStr D, T, Drs
InsTblzDy D, T, Drs.Dy
End Sub

Sub CrtTzDrszEmpzAllStr(D As Database, T, Drs As Drs)
Dim C&, F, O$(), Dy()
Dy = Drs.Dy
For Each F In Drs.Fny
    If IsMemCol(ColzDy(Dy, C)) Then
        PushI O, "M:" & F
    Else
        PushI O, F
    End If
    C = C + 1
Next
CrtTzShtTyscfBql D, T, Jn(O, "`")
End Sub

Sub CrtTblOfEmpzDrs(D As Database, T, Drs As Drs)
CrtTzShtTyscfBql D, T, ShtTyscfBqlzDrs(Drs)
End Sub

Private Sub Z_ShtTyscfBqlzDrs()
Dim Drs As Drs
GoSub T0
Exit Sub
T0:
    Drs = SampDrs
    Ept = "A`B:B`Byt:C`I:D`L:E`D:G`S:H`C:I`Dte:J`M:K"
    GoTo Tst
Tst:
    Act = ShtTyscfBqlzDrs(Drs)
    C
    Return
End Sub

Function ShtTyszCol$(Col())
Dim O$
Select Case True
Case IsBoolCol(Col): O = "B"
Case IsDteCol(Col): O = "Dte"
Case IsColOfNum(Col): O = ShtTyzNumCol(Col)
Case IsColOfStr(Col): O = IIf(IsMemCol(Col), "M", "")
Case Else: Thw CSub, "Col cannot determine its type: Not [Str* Num* Bool* Dte*:Col]", "Col", Col
End Select
ShtTyszCol = O
End Function
Function ShtTyzNumCol$(Col)
ShtTyzNumCol = ShtTyzDao(DaoTyzNumCol(Col))
End Function
Function IsMemCol(Col) As Boolean
Dim I
For Each I In Col
    If IsStr(I) Then
        If Len(I) > 255 Then IsMemCol = True: Exit Function
    End If
Next
End Function

Function IsColOfStr(Col) As Boolean
Dim V
For Each V In Col
    If Not IsStr(V) Then Exit Function
Next
IsColOfStr = True
End Function

Function DaoTyzNumCol$(NumCol)
ThwIf_NotAy NumCol, CSub
Dim O As VbVarType: O = VarType(NumCol(0))
If Not IsNumzVbTy(O) Then Stop
Dim V
For Each V In NumCol
    O = MaxNumVbTy(O, VarType(V))
Next
DaoTyzNumCol = DaoTyzVbTy(O)
End Function
Function IsColOfNum(Col) As Boolean
Dim V
For Each V In Col
    If Not IsNumeric(V) Then Exit Function
Next
IsColOfNum = True
End Function

Function IsNumzVbTy(A As VbVarType) As Boolean
Select Case A
Case vbByte, vbInteger, vbLong, vbSingle, vbDecimal, vbDouble, vbCurrency: IsNumzVbTy = True
End Select
End Function

Private Function MaxNumVbTy(A As VbVarType, B As VbVarType) As VbVarType
If A = B Then MaxNumVbTy = A: Exit Function
If Not IsNumzVbTy(B) Then Thw CSub, "Given B is not NumVbTy", "B-VarType", B
Dim O As VbVarType
Select Case A
Case VbVarType.vbByte:      O = B
Case VbVarType.vbInteger:   O = IIf(B = vbByte, A, B)
Case VbVarType.vbLong:      O = IIf((B = vbByte) Or (B = vbInteger), A, B)
Case VbVarType.vbSingle:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong), A, B)
Case VbVarType.vbDecimal:   O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle), A, B)
Case VbVarType.vbDouble:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal), A, B)
Case VbVarType.vbCurrency:  O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal) Or (B = vbDouble), A, B)
Case Else:                  Thw CSub, "Given A is not NumVbTy", "A-VarType", A
End Select
MaxNumVbTy = O
End Function

Function ShtTyzNumVbTy$(NumVbTy As VbVarType)
Dim O$
Select Case NumVbTy
Case VbVarType.vbByte:      O = "Byt:"
Case VbVarType.vbCurrency:  O = "C:"
Case VbVarType.vbDecimal:   O = "Dec:"
Case VbVarType.vbDouble:    O = "D:"
Case VbVarType.vbInteger:   O = "I:"
Case VbVarType.vbLong:      O = "L:"
Case VbVarType.vbSingle:    O = "S:"
Case Else: Thw CSub, "NumVbTy is not numeric VbTy", "NumVbTyp", ShtTyzNumVbTy(NumVbTy)
End Select
End Function

Function IsBoolCol(Col()) As Boolean
Dim V
For Each V In Col
    If Not IsBool(V) Then Exit Function
Next
IsBoolCol = True
End Function

Function IsDteCol(Col()) As Boolean
Dim V
For Each V In Col
    If Not IsDte(V) Then Exit Function
Next
IsDteCol = True
End Function

Sub InsTblzDy(D As Database, T, Dy())
InsRszDy RszT(D, T), Dy
End Sub

Sub CrtTzJnFld(D As Database, T, KK$, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
Rq D, SqlSel_Fny_Into_T_OB(Ny(KK), Tar, T)
AddFld D, T, LisFld, dbMemo
Dim KKIdx&(), JnFldIx&
    KKIdx = Ixy(Fny(D, T), Ny(KK))
    JnFldIx = IxzTF(D, T, JnFld)
InsTblzDy D, T, DyoJnFldKK(DyoT(D, T), KKIdx, JnFldIx)
End Sub

Function IxzTF%(D As Database, T, Fld$)
Dim F As Dao.Field, O%
For Each F In D.TableDefs(T).Fields
    If F.Name = Fld Then
        IxzTF = O
        Exit Function
    End If
    O = O + 1
Next
IxzTF = -1
End Function
Sub CrtPk(D As Database, T)
D.Execute SqlCrtPk_T(T)
End Sub

Function JnQSqCommaSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnQSqCommaSpcAp = JnQSqCommaSpc(SyzAy(Av))
End Function
Function CommaSpcSqAv$(Av())

End Function
Function JnCommaSpcFF$(FF$)
JnCommaSpcFF = JnQSqCommaSpc(TermAy(FF))
End Function

Sub CrtSk(D As Database, T, Skff$)
D.Execute SqlCrtSk_T_SkFF(T, Skff)
End Sub

Sub DrpFld(D As Database, T, FF$)
Dim F$, I
For Each I In ItrzTT(FF)
    F = I
    D.Execute SqlDrpCol_T_F(T, F)
Next
End Sub

Sub RenFld(D As Database, T, F$, ToFld$)
D.TableDefs(T).Fields(F).Name = ToFld
End Sub

Function FdStrzTF$(D As Database, T, F$)
FdStrzTF = FdStr(FdzTF(D, T, F$))
End Function

Function IntAyzDbtf(D As Database, T, F$) As Integer()
Q = FmtQQ("Select [?] from [?]", F, T)
IntAyzDbtf = IntAyzQ(D, Q)
End Function

Function NxtId&(D As Database, T)
Dim S$: S = FmtQQ("select Max(?Id) from [?]", T, T)
NxtId = VzQ(D, S) + 1
End Function

Function DaoTyzTF(D As Database, T, F) As Dao.DataTypeEnum
DaoTyzTF = D.TableDefs(T).Fields(F).Type
End Function

Function ShtTyzTF$(D As Database, T, F$)
ShtTyzTF = ShtTyzDao(DaoTyzTF(D, T, F$))
End Function

Function CnStrzLnkTbl$(D As Database, T)
CnStrzLnkTbl = D.TableDefs(T).Connect
End Function

Sub AddFldzExpr(D As Database, T, F$, Expr$, Ty As Dao.DataTypeEnum)
D.TableDefs(T).Fields.Append Fd(F, Ty, Expr:=Expr)
End Sub

Function VzTFRecId(D As Database, T, F$, RecId&) ' K is Pk value
'VzTFRecId = VzQ(D, SqlSel_FF_T(F, T, BexpRecId(T, RecId)))
End Function

Sub CrtTzCloneEmp(D As Database, T, FmTbl$)
Rq D, SqlSel_Into_T_WhFalse(T, FmTbl)
End Sub

Sub KillIfTmpDb(D As Database)
If IsDbTmp(D) Then
    Dim Fb$: Fb = D.Name
    ClsDb D
    Kill Fb
End If
End Sub

Private Sub Z_CrtDupKeyTbl()
Dim D As Database: Set D = TmpDb
DrpTT D, "#A #B"
'T = "AA"
CrtTzDup D, "#A", "#B", "Sku BchNo"
DrpDbIfTmp D
End Sub

Private Sub Z_PkFny()
Z:
    Dim D As Database
    Set D = Db(SampFbzDutyDta)
    Dim Dr(), Dy(), T, I
    For Each I In Tny(D)
        T = I
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFny(D, T)
        PushI Dy, Dr
    Next
    BrwDy Dy
    Exit Sub
End Sub

Private Sub Z()
Dim Db As Database
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
Dim O As Database
Dim XX
End Sub

Function VzArs(A As AdoDb.Recordset)
If NoReczAdo(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
VzArs = V
End Function

Function VzCnq(A As AdoDb.Connection, Q)
VzCnq = VzArs(A.Execute(Q))
End Function

Function NReczFxw&(Fx, Wsn, Optional Bexp$)
NReczFxw = VzCnq(CnzFx(Fx), SqlSelCnt_T_OB(CatTnzWsn(Wsn), Bexp))
End Function
Function NReczT&(D As Database, T, Optional Bexp$)
NReczT = VzQ(D, SqlSelCnt_T_OB(T, Bexp))
End Function

Property Get LofVblzDbt$(D As Database, T)
LofVblzDbt = TblPrp(D, T, "LofVbl")
End Property

Property Let LofVblzDbt(D As Database, T, LofVbl$)
TblPrp(D, T, "LofVbl") = LofVbl
End Property

Function IsLnk(D As Database, T) As Boolean
IsLnk = IsLnkzFb(D, T) Or IsLnkzFx(D, T)
End Function

Function CnStrzT$(D As Database, T)
On Error Resume Next
CnStrzT = D.TableDefs(T).Connect
End Function

Function IsLnkzFb(D As Database, T) As Boolean
IsLnkzFb = HasPfx(CnStrzT(D, T), ";Database=")
End Function

Function IsLnkzFx(D As Database, T) As Boolean
IsLnkzFx = HasPfx(CnStrzT(D, T), "Excel")
End Function

