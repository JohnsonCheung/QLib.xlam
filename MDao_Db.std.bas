Attribute VB_Name = "MDao_Db"
Option Explicit
Public CDb As Database
Public Q$
Sub OpnCDb(Fb)
Set CDb = Db(Fb)
End Sub

Property Get IsOkDb(A As Database) As Boolean
On Error GoTo X
IsOkDb = A.Name = A.Name
Exit Property
X:
End Property

Function AddTd(A As Database, Td As DAO.TableDef) As DAO.TableDef
A.TableDefs.Append Td
Set AddTd = Td
End Function

Sub AddTmpTbl(A As Database)
A.TableDefs.Append TmpTd
End Sub

Function PthzDb$(A As Database)
PthzDb = Pth(DbNm(A))
End Function
Function IsTmpDbz(A As Database) As Boolean
IsTmpDbz = PthzDb(A) = TmpDbPth
End Function
Sub DrpDbIfTmp()
DrpDbIfTmpz CDb
End Sub

Sub DrpDbIfTmpz(A As Database)
If IsTmpDbz(A) Then
    Dim N$
    N = DbNm(A)
    A.Close
    DltFfn N
End If
End Sub

Sub BrwDb(A As Database)
BrwFb A.Name
End Sub

Function Stru(TT) As String()
End Function

Function TnyzTT(TT) As String()
TnyzTT = TermAy(TT)
End Function

Function StruzTT(A As Database, TT)
Dim T
For Each T In Itr(AySrt(TnyzTT(TT)))
    PushI StruzTT, StruzT(A, T)
Next
End Function

Function Struz(Db As Database) As String()
'Struz = StruzTT(Db, TnyzDaoDb(Db))
End Function

Function OupTny() As String()
OupTny = OupTnyz(CDb)
End Function

Function OupTnyz(A As Database) As String()
OupTnyz = AywPfx(Tnyz(A), "@")
End Function
Sub CrtTblzShtTySemiFldSslDb(A As Database, T, ShtTySemiFldSsl$)
A.TableDefs.Append TdTblShtTySemiFldSsl(T, ShtTySemiFldSsl)
End Sub
Sub CrtTblzShtTySemiFldSsl(T, ShtTySemiFldSsl$)
CrtTblzShtTySemiFldSslDb CDb, T, ShtTySemiFldSsl
End Sub

Sub Drpz(A As Database, TT, Optional NoReOpn As Boolean)
Dim T
For Each T In ItrTT(TT)
    DrpzT Dbz(A, NoReOpn), T
Next
End Sub

Sub Drp(TT)
Drpz CDb, TT
End Sub

Sub DrpzT(A As Database, T)
If HasTblz(A, T) Then A.TableDefs.Delete T
End Sub

Sub CrtTblz(A As Database, T$, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function Dsz(A As Database, Optional DsNm$) As Ds
Dim Nm$
If DsNm = "" Then
    Nm = DbNm(A)
Else
    Nm = DsNm
End If
Set Dsz = DszTT(A, Tnyz(A), Nm)
End Function

Function DszTT(A As Database, TT, Optional DsNm$) As Ds
Dim DtAy() As Dt
    Dim U%, Tny$()
    Tny = DftTny(TT, A.Name)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        'Set DtAy(J) = Dt(A, Tny(J))
    Next
'Set DsDb = Ds(DtAy, DftDbNm(DsNm, A))
End Function
Sub EnsTmpTblz(A As Database)
If HasTblz(A, "#Tmp") Then Exit Sub
A.Execute "Create Table [#Tmp] (AA Int, BB Text 10)"
End Sub

Sub EnsTmpTbl()
EnsTmpTblz CDb
End Sub
Sub RunQz(A As Database, Q)
On Error GoTo X
A.Execute Q
Exit Sub
X: Dim E$: E = Err.Description: Thw CSub, "Running Sql error", "Er Sql Db", E, Q, DbNm(A)
End Sub

Sub RunQQz(A As Database, QQ, ParamArray Ap())
Dim Av(): Av = Ap
RunQz A, FmtQQAv(QQ, Av)
End Sub

Sub RunQQ(QQ, ParamArray Ap())
Dim Av(): Av = Ap
RunQ FmtQQAv(QQ, Av)
End Sub

Sub RunQ(Q)
RunQz CDb, Q
End Sub

Function RszQQ(QQ, ParamArray Ap()) As DAO.Recordset
Dim Av(): Av = Ap
Set RszQQ = Rs(FmtQQAv(QQ, Av))
End Function
Function Rs(Q) As DAO.Recordset
Set Rs = Rsz(CDb, Q)
End Function
Function RszT(A As Database, T) As DAO.Recordset
On Error GoTo X
Set RszT = A.TableDefs(T).OpenRecordset
Exit Function
X: Thw CSub, "Error in opening Table", "Er T Db", Err.Description, T, DbNm(A)
End Function

Function Rsz(A As Database, Q) As DAO.Recordset
On Error GoTo X
Set Rsz = A.OpenRecordset(Q)
Exit Function
X: Thw CSub, "Error in opening Rs", "Er Sql Db", Err.Description, Q, DbNm(A)
End Function

Function HasReczDbq(A As Database, Q) As Boolean
HasReczDbq = HasReczRs(Rsz(A, Q))
End Function

Function HasQryz(A As Database, Q) As Boolean
HasQryz = HasReczDbq(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function HasTblz(A As Database, T, Optional ReOpn As Boolean)
HasTblz = HasItn(Dbz(A, ReOpn).TableDefs, T)
End Function

Function HasTblzMSysObjDb(A As Database, T) As Boolean
HasTblzMSysObjDb = HasReczRs(Rsz(A, FmtQQ("Select Name from MSysObjects where Type in (1,6) and Name='?'", T)))
End Function

Function IsDbOkz(A As Database) As Boolean
On Error GoTo X
IsDbOkz = IsStr(CDb.Name)
Exit Function
X:
End Function
Function IsDbOk() As Boolean
IsDbOk = IsDbOkz(CDb)
End Function
Function DbPth$()
DbPth = DbPthz(CDb)
End Function
Function DbPthz$(A As Database)
DbPthz = FfnPth(DbNm(A))
End Function

Function Qnyz(A As Database) As String()
Qnyz = SyzDbq(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function
Function Qny() As String()
Qny = Qnyz(CDb)
End Function

Function Rsry(A As Database, Qry) As DAO.Recordset
Set Rsry = A.QueryDefs(Qry).OpenRecordset
End Function

Function SrcTnAyz(A As Database) As String()
Dim T
For Each T In Tni
    PushNonBlankStr SrcTnAyz, SrcTz(A, T)
Next
End Function
Function SrcTnAy() As String()
SrcTnAy = SrcTnAyz(CDb)
End Function

Function TmpTny() As String()
TmpTny = AywPfx(Tny, "#")
End Function

Function Tniz(A As Database)
Asg Itr(Tnyz(A)), Tniz
End Function

Function Tni()
Asg Tniz(CDb), Tni
End Function

Function Tnyz(A As Database) As String()
Set A = DAO.DBEngine.OpenDatabase(A.Name)
Dim T As TableDef
For Each T In A.TableDefs
    If Not IsSysTd(T) Then
        If Not IsHidTd(T) Then
            PushI Tny, T.Name
        End If
    End If
Next
End Function

Function Tny() As String()
Tny = Tnyz(CDb)
End Function

Function TnyzADO(A As Database) As String()
TnyzADO = TnyzFb(A.Name)
End Function
Function TnyzDaoFb(Fb) As String()
TnyzDaoFb = TnyzDaoDb(Db(Fb))
End Function
Function TnyzDaoDb(A As Database, Optional NoReOpn As Boolean) As String()
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In Dbz(A, NoReOpn).TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI TnyzDaoDb, T.Name
    End Select
Next
End Function

Function TnyzMSysObj() As String()
TnyzMSysObj = SyzQ("Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
End Function

Private Sub ZZ_Qny()
'DmpAy Qny(Db(SampFbzDutyDta))
End Sub

Private Sub Z_Ds()
Dim Db As Database, Tny0
Stop
ZZ1:
    Set Db = Db(SampFbzDutyDta)
'    Set Act = Ds(Db)
    CvDs(Act).Brw
    Exit Sub
ZZ2:
    Tny0 = "Permit PermitD"
    'Set Act = Ds( Tny0)
    Stop
End Sub

Private Sub Z_Qny()
'DmpAy Qny(CDb)
End Sub

Private Sub ZZ()
Dim A As Database
Dim B As DAO.TableDef
Dim C$()
Dim D As Variant
Dim E$
Dim F As Drs
Dim G As Dictionary
'AddTd A, B
'ChkPk A, C
'ChkSk A, C
'DbCnSy A
'TblDesAy A
'Db_Drp_Qry A, D
'DsDb A, D, E
'EnsTmp1TblDb A
'HasQry A, D
'HasDbt A, D
End Sub

Sub RenTblzAddPfx(TT, Pfx$)
RenTblzAddPfxDb CDb, TT, Pfx
End Sub
Sub RenTblzAddPfxDb(A As Database, TT, Pfx$)
Dim T
For Each T In ItrTT(TT)
    RenTblzAddPfxDbt A, T, Pfx
Next
End Sub

Sub BrwTblz(A As Access.Application, T)
A.DoCmd.OpenTable T
End Sub

Sub BrwTTzAcs(A As Access.Application, TT)
Dim T
For Each T In ItrTT(TT)
    BrwTblz A, T
Next
End Sub

Sub BrwTTz(A As Database, TT)
BrwTTzAcs CAcs, TT
End Sub

Sub BrwTT(TT)
BrwTTz CDb, TT
End Sub

Function TdStrAyz(A As Database, TT) As String()
Dim T
For Each T In ItrTT(TT)
    PushI TdStrAyz, TdStrz(A, T)
Next
End Function
Function TdStrAy(TT) As String()
TdStrAy = TdStrAyz(CDb, TT)
End Function

Sub CrtTmpTblz(A As Database)
Drp "#Tmp"
A.TableDefs.Append TmpTd
End Sub
Sub RenTblz(A As Database, T, ToNm$)
A.TableDefs(T).Name = ToNm
End Sub
Sub RenTbl(T, ToNm$)
RenTblz CDb, T, ToNm
End Sub

Sub RenTblzFmPfx(FmPfx$, ToPfx$)
Dim T As TableDef
For Each T In CDb.TableDefs
    If HasPfx(T.Name, FmPfx) Then
        T.Name = RplPfx(T.Name, FmPfx, ToPfx)
    End If
Next
End Sub

Sub DrpzTT(TT)
End Sub

Sub DrpzAp(ParamArray TblAp())
Dim Av(): Av = TblAp
DrpzTT Av
End Sub

Property Get TblDesz$(A As Database, T)

End Property

Property Let TblDesz(A As Database, T, Des$)

End Property

Property Get TblAttDesz$()
TblAttDesz = TblDesz(CDb, "Att")
End Property

Property Let TblAttDesz(Des$)
TblDesz(CDb, "Att") = Des
End Property

Property Get TblAttDes$()
TblAttDes = TblAttDesz()
End Property

Property Let TblAttDes(Des$)
TblAttDesz = Des
End Property

Function ValQ(Q)
ValQ = ValzDbq(CDb, Q)
End Function

Property Set TblDesDic(D As Dictionary)
Dim I, F$, T$
For Each I In D.Keys
    AsgBrkDot1 I, T, F
    FldDesz(CDb, T, F) = D(I)
Next
End Property

Property Get FldDesDicz(A As Database) As Dictionary
Dim T, F, D$
Set FldDesDic = New Dictionary
For Each T In Tniz(A)
    For Each F In Fnyz(A, T)
        D = FldDesz(A, T, F)
        If D <> "" Then FldDesDicz.Add T & "." & F, D
    Next
Next
End Property
Property Set FldDesDicz(A As Database, D As Dictionary)
Dim T$, F$, Des$, TDotF
Set FldDesDic = New Dictionary
For Each TDotF In D.Keys
    AsgBrkDot TDotF, T, F
    Des = D(TDotF)
    FldDesz(A, T, F) = Des
Next
End Property

Property Get FldDesDic() As Dictionary
Set FldDesDic = FldDesDicz(CDb)
End Property

Property Set FldDesDic(D As Dictionary)
Set FldDesDicz(CDb) = D
End Property

Property Get TblDesDic() As Dictionary
Set TblDesDic = TblDesDicz(CDb)
End Property
Property Set TblDesDicz(A As Database, TblDic As Dictionary)

End Property

Property Get TblDesDicz(A As Database) As Dictionary
Dim T, O As New Dictionary
For Each T In Tniz(A)
    AddDiczNonBlankStr O, T, TblDesz(A, T)
Next
Set TblDesDicz = O
End Property


Function JnStrDiczTwoColSql(TwoColSql) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
'Set JnStrDiczTwoColSql = JnStrDicTwoFldRs(RszSql(TwoColSql))
End Function


Sub AppTdAyz(A As Database, TdAy() As DAO.TableDef)
Dim T
For Each T In Itr(TdAy)
    A.TableDefs.Append T
Next
End Sub

Private Sub ZZ_BrwTbl()
Drp "#A #B"
RunQ "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_UpdSeqFld"
'BrwTblAp "#A", "Sku BchNo", "#B"
Drp "#B"
End Sub

Function TnizInp(A As Database)
Asg Itr(TnyzInp(A)), TnizInp
End Function

Function TnyzInp(A As Database) As String()
TnyzInp = AywLik(Tnyz(A), ">*")
End Function

Sub CpyInpTblAsTmpz(A As Database)
Dim T
For Each T In TnizInp(A)
    Drp "#I" & T
    RunQQ "Select * into [#I?] from [?]", T, T
Next
End Sub

Function Dbz(A As Database, Optional NoReOpn As Boolean) As Database
If NoReOpn Then Set Dbz = A Else Set Dbz = Db(A.Name)
End Function

Function DbNm$(A As Database)
On Error GoTo X
DbNm = A.Name
Exit Function
X:
DbNm = Err.Description
End Function

