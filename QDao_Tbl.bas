Attribute VB_Name = "QDao_Tbl"
Option Explicit
Private Const CMod$ = "MDao_Tbl."
Private Const Asm$ = "QDao"
Const C_SkNm$ = "SecondaryKey"

Function FnyzT(A As Database, T) As String()
FnyzT = Fny(A, T)
End Function

Function Fny(A As Database, T) As String()
Fny = Itn(A.TableDefs(T).Fields)
End Function

Function ColzRs(A As Database, T, F$) As Dao.Recordset
Set ColzRs = Rs(A, SqlSel_F_T(F, T))
End Function

Function ColSetzT(A As Database, T, F$) As Aset
Set ColSetzT = ColSetzRs(ColzRs(A, T, F))
End Function

Function CntDiczTF(A As Database, T, F$) As Dictionary
Set CntDiczTF = CntDiczRs(ColzRs(A, T, F$))
End Function

Function IdxzTd(A As Dao.TableDef, IdxNm$) As Dao.Index
Set IdxzTd = FstItmzNm(A.Indexes, IdxNm$)
End Function

Function Idx(A As Database, T, IdxNm$) As Dao.Index
Set Idx = IdxzTd(A.TableDefs(T), IdxNm)
End Function

Function HasSk(A As Database, T) As Boolean
HasSk = Not IsNothing(SkIdx(A, T))
End Function

Function HasIdx(A As Database, T, IdxNm$) As Boolean
HasIdx = HasItn(A.TableDefs(T).Indexes, IdxNm)
End Function

Function FstUniqIdx(A As Database, T) As Dao.Index
Set FstUniqIdx = FstItmTrueP(A.TableDefs(T).Indexes, PrpPth("Unique"))
End Function

Function HasFld(A As Database, T, F$) As Boolean
HasFld = HasItn(A.TableDefs(T).Fields, F)
End Function

Function HasPk(A As Database, T) As Boolean
HasPk = HasPkzTd(A.TableDefs(T))
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

Function HasStdPk(A As Database, T) As Boolean
HasStdPk = HasStdPkzTd(A.TableDefs(T))
End Function

Function HasId(A As Database, T, Id&) As Boolean
If HasPk(A, T) Then
    HasId = HasRec(RszId(A, T, Id))
End If
End Function

Function DryzTFF(A As Database, T, FF$) As Variant()
DryzTFF = DryzQ(A, SqlSel_FF_T(FF, T))
End Function

Sub AsgColApzDrsFF(A As Drs, FF$, ParamArray OColAp())
Dim F, J%
For Each F In TermAy(FF)
    OColAp(J) = ColzDrs(A, CStr(F))
    J = J + 1
Next
End Sub

Function RszId(A As Database, T, Id&) As Dao.Recordset
Set RszId = Rs(A, SqlSel_T_WhId(T, Id))
End Function

Function CsvLyzDbt(A As Database, T) As String()
CsvLyzDbt = CsvLyzRs(RszT(A, T))
End Function

Function DrszT(A As Database, T) As Drs
DrszT = DrszRs(RszT(A, T))
End Function
Function DryzT(A As Database, T) As Variant()
DryzT = DryzRs(RszT(A, T))
End Function

Function DtzT(A As Database, T) As Dt
DtzT = Dt(T, Fny(A, T), DryzT(A, T))
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
SrcFbzT = Bet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function NColzT&(A As Database, T)
NColzT = A.TableDefs(T).Fields.Count
End Function

Function NReczDbtBexpr&(A As Database, T, Bexpr$)
NReczDbtBexpr = ValzQ(A, FmtQQ("Select Count(*) from [?]?", T, SqpWh(Bexpr)))
End Function

Function PkFnyzTd(A As Dao.TableDef) As String()
PkFnyzTd = FnyzIdx(PkizTd(A))
End Function

Function PkFny(A As Database, T) As String()
PkFny = FnyzIdx(PkIdx(A, T))
End Function

Function PkIdxNm$(A As Database, T)
PkIdxNm = ObjNm(PkIdx(A, T))
End Function

Function PkizTd(A As Dao.TableDef) As Dao.Index
Set PkizTd = FstItn(A.Indexes, C_PkNm)
End Function

Function PkIdx(A As Database, T) As Dao.Index
Set PkIdx = PkizTd(A.TableDefs(T))
End Function

Function RszTFny(A As Database, T, Fny$()) As Dao.Recordset
Set RszTFny = A.OpenRecordset(SqlSel_Fny_T(Fny, T))
End Function

Function RszTFF(A As Database, T, FF$) As Dao.Recordset
Set RszTFF = RszTFny(A, T, Ny(FF))
End Function

Function RszTF(A As Database, T, F$) As Dao.Recordset
Set RszTF = A.OpenRecordset(SqlSel_F_T(F, T))
End Function

Function RszT(A As Database, T) As Dao.Recordset
Set RszT = A.TableDefs(T).OpenRecordset
End Function

Function FdzTF(A As Database, T, F$) As Dao.Field2
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
'If Not Has Then FunMsgAp_Dmp_Ly CSub, "[Db] has not such [Tbl]", Dbn, T: Exit Function
Dim F$()
    F = Fny(A, T)
    If IsLnkzFx(A, T) Then
        StruzT = T & " " & JnSpc(QuoteSqzAyIf(F))
        Exit Function
    End If

Dim P$
    If HasEle(F, T & "Id") Then
        P = " *Id"
        F = MinusAy(F, Array(T & "Id"))
    End If
Dim Sk$()
    Sk = SkFny(A, T)

Dim R$
    Dim I
    Dim Rst$()
    Rst = RplStarzSy(CvSy(MinusAy(F, Sk)), T)
    R = " " & JnSpc(QuoteSqzAyIf(Rst))

Dim S$
    S = JnSpc(QuoteSqzAyIf(RplStarzSy(Sk, T)))
    If S <> "" Then S = " " & S & " |"

StruzT = T & P & S & R
End Function

Function LasUpdTim(A As Database, T) As Date
LasUpdTim = TblPrp(A, T, "LastUpdated")
End Function

Sub InsTblzDrs(A As Database, T, B As Drs)
InsRszDry RszTFny(A, T, B.Fny), B.Dry
End Sub

Sub AddFd(A As Database, T, Fd As Dao.Fields)
A.TableDefs(T).Fields.Append Fd
End Sub

Sub AddFld(A As Database, T, F$, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
If HasFld(A, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = SqlTyzDao(Ty, Si, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
A.Execute S
End Sub

Sub RenTblzAddPfx(A As Database, T, Pfx$)
RenTbl A, T, Pfx & T
End Sub

Sub BrwTblzByDt(A As Database, T)
BrwDt DtzT(A, T)
End Sub

Function IsSysTbl(A As Database, T) As Boolean
IsSysTbl = (A.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbSystemObject) <> 0
End Function

Function IsHidTbl(A As Database, T) As Boolean
IsHidTbl = (A.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbHiddenObject) <> 0
End Function

Function Lnkinf(A As Database) As String()
Dim T$, I
For Each I In Tni(A)
    T = I
    PushI Lnkinf, LnkinfzT(A, T)
Next
End Function

Function LnkinfzT$(A As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(A, T): LnkinfzT = FmtQQ("LnkFx(?).LnkWs(?).Tbl(?).Db(?)", CnStrzDbt(A, T), SrcTn(A, T), T, Dbn(A))
Case IsLnkzFb(A, T): LnkinfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", CnStrzDbt(A, T), SrcTn(A, T), T, Dbn(A))
End Select
End Function

Function CrtTblzDupKey$(A As Database, Into$, FmTbl, KK$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SyzSS(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, FmTbl, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", Into$, FmTbl, Tmp, Jn)
DrpT A, Tmp
End Function

Private Sub Z_CrtTblzDrs()
Dim D As Database
GoSub ZZ
Exit Sub
ZZ:
    Set D = TmpDb
    DrpTmp D
    CrtTblzDrs D, "#A", SampDrs
    BrwDb D
    Return
End Sub
Sub CrtTblzDrs(A As Database, T, Drs As Drs)
CrtTblOfEmpzDrs A, T, Drs
InsTblzDry A, T, Drs.Dry
End Sub

Sub CrtTblzDrszAllStr(A As Database, T, Drs As Drs)
CrtTblzDrszEmpzAllStr A, T, Drs
InsTblzDry A, T, Drs.Dry
End Sub

Sub CrtTblzDrszEmpzAllStr(A As Database, T, Drs As Drs)
Dim C&, F, O$(), Dry()
Dry = Drs.Dry
For Each F In Drs.Fny
    If IsMemCol(ColzDry(Dry, C)) Then
        PushI O, "M:" & F
    Else
        PushI O, F
    End If
    C = C + 1
Next
CrtTblzShtTyscfBql A, T, Jn(O, "`")
End Sub

Sub CrtTblOfEmpzDrs(A As Database, T, Drs As Drs)
CrtTblzShtTyscfBql A, T, ShtTyscfBqlzDrs(Drs)
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

Sub InsTblzDry(A As Database, T, Dry())
InsRszDry RszT(A, T), Dry
End Sub

Sub CrtTblzJnFld(A As Database, T, KK$, JnFld$, Optional Sep$ = " ")
Dim Tar$, LisFld$
    Tar = T & "_Jn_" & JnFld
    LisFld = JnFld & "_Jn"
RunQ A, SqlSel_Fny_Into_T(Ny(KK), Tar, T)
AddFld A, T, LisFld, dbMemo
Dim KKIdx&(), JnFldIx&
    KKIdx = Ixy(Fny(A, T), Ny(KK))
    JnFldIx = IxzTF(A, T, JnFld)
InsTblzDry A, T, DryzJnFldKK(DryzT(A, T), KKIdx, JnFldIx)
End Sub

Function IxzTF%(A As Database, T, Fld$)
Dim F As Dao.Field, O%
For Each F In A.TableDefs(T).Fields
    If F.Name = Fld Then
        IxzTF = O
        Exit Function
    End If
    O = O + 1
Next
IxzTF = -1
End Function
Sub CrtPk(A As Database, T)
A.Execute SqlCrtPkzT(T)
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

Sub CrtSk(A As Database, T, Skff$)
A.Execute SqlCrtSk_T_SkFF(T, Skff)
End Sub

Sub DrpFld(A As Database, T, FF$)
Dim F$, I
For Each I In ItrzTT(FF)
    F = I
    A.Execute SqlDrpCol_T_F(T, F)
Next
End Sub

Sub RenFld(A As Database, T, F$, ToFld$)
A.TableDefs(T).Fields(F).Name = ToFld
End Sub

Function FdStrzTF$(A As Database, T, F$)
FdStrzTF = FdStr(FdzTF(A, T, F$))
End Function

Function IntAyzDbtf(A As Database, T, F$) As Integer()
Q = FmtQQ("Select [?] from [?]", F, T)
IntAyzDbtf = IntAyzQ(A, Q)
End Function

Function NxtId&(A As Database, T)
Dim S$: S = FmtQQ("select Max(?Id) from [?]", T, T)
NxtId = ValzQ(A, S) + 1
End Function

Function DaoTyzTF(A As Database, T, F) As Dao.DataTypeEnum
DaoTyzTF = A.TableDefs(T).Fields(F).Type
End Function

Function ShtTyzTF$(A As Database, T, F$)
ShtTyzTF = ShtTyzDao(DaoTyzTF(A, T, F$))
End Function

Function CnStrzLnkTbl$(A As Database, T)
CnStrzLnkTbl = A.TableDefs(T).Connect
End Function

Sub AddFldzExpr(A As Database, T, F$, Expr$, Ty As Dao.DataTypeEnum)
A.TableDefs(T).Fields.Append Fd(F, Ty, Expr:=Expr)
End Sub

Function ValzTFRecId(A As Database, T, F$, RecId&) ' K is Pk value
ValzTFRecId = ValzQ(A, SqlSel_FF_T(F, T, BexprRecId(T, RecId)))
End Function

Sub CrtTblzCloneEmp(A As Database, T, FmTbl$)
RunQ A, SqlSel_Into_T_WhFalse(T, FmTbl)
End Sub

Sub KillIfTmpDb(A As Database)
If IsTmpDb(A) Then
    Dim Fb$: Fb = A.Name
    ClsDb A
    Kill Fb
End If
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
    Dim Dr(), Dry(), T, I
    For Each I In Tny(A)
        T = I
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFny(A, T)
        PushI Dry, Dr
    Next
    BrwDry Dry
    Exit Sub
End Sub

Private Sub ZZ()
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
Dim O As Dao.Database
Dim XX
End Sub

Function ValzArs(A As AdoDb.Recordset)
If NoReczAdo(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
ValzArs = V
End Function

Function ValzCnq(A As AdoDb.Connection, Q)
ValzCnq = ValzArs(A.Execute(Q))
End Function

Function NReczFxw&(Fx, Wsn, Optional Bexpr$)
NReczFxw = ValzCnq(CnzFx(Fx), SqlSelCnt_T(CatTn(Wsn), Bexpr))
End Function
Function NReczT&(A As Database, T, Optional Bexpr$)
NReczT = ValzQ(A, SqlSelCnt_T_OB(T, Bexpr))
End Function

Property Get LofVblzDbt$(A As Database, T)
LofVblzDbt = TblPrp(A, T, "LofVbl")
End Property

Property Let LofVblzDbt(A As Database, T, LofVbl$)
TblPrp(A, T, "LofVbl") = LofVbl
End Property

Function IsLnk(A As Database, T) As Boolean
IsLnk = IsLnkzFb(A, T) Or IsLnkzFx(A, T)
End Function

Function CnStrzT$(A As Database, T)
On Error Resume Next
CnStrzT = A.TableDefs(T).Connect
End Function

Function IsLnkzFb(A As Database, T) As Boolean
IsLnkzFb = HasPfx(CnStrzT(A, T), ";Database=")
End Function

Function IsLnkzFx(A As Database, T) As Boolean
IsLnkzFx = HasPfx(CnStrzT(A, T), "Excel")
End Function

