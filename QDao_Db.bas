Attribute VB_Name = "QDao_Db"
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Db."
Public Q$

Function IsOkDb(A As Database) As Boolean
On Error GoTo X
IsOkDb = A.Name = A.Name
Exit Function
X:
End Function

Sub AddTmpTbl(A As Database)
A.TableDefs.Append TmpTd
End Sub

Function PthzDb$(A As Database)
PthzDb = Pth(Dbn(A))
End Function
Function IsTmpDb(A As Database) As Boolean
IsTmpDb = PthzDb(A) = TmpPthzDb
End Function


Sub DrpDbIfTmp(A As Database)
If IsTmpDb(A) Then
    Dim N$
    N = Dbn(A)
    A.Close
    DltFfn N
End If
End Sub

Sub BrwDb(A As Database)
BrwFb A.Name
End Sub

Function StruzTny(A As Database, Tny$())
Dim I
For Each I In Itr(QSrt(Tny))
    PushI StruzTny, StruzT(A, CStr(I))
Next
End Function

Function StruzTT(A As Database, TT$)
StruzTT = StruzTny(A, Ny(TT))
End Function

Function Stru(A As Database) As String()
Stru = StruzTny(A, Tny(A))
End Function

Function OupTny(A As Database) As String()
OupTny = AywPfx(Tny(A), "@")
End Function
Sub DrpTny(A As Database, Tny$())
Dim T
For Each T In Tny
    DrpT A, CStr(T)
Next
End Sub
Sub DrpTT(A As Database, TT$)
Dim T
End Sub
Sub DrpTmp(A As Database)
DrpTny A, TmpTny(A)
End Sub
Sub DrpT(A As Database, T)
If HasTbl(A, T) Then A.TableDefs.Delete T
End Sub

Sub CrtTbl(A As Database, T, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DszDb(A As Database, Optional DsNm$) As Ds
Dim Nm$
If DsNm = "" Then
    Nm = Dbn(A)
Else
    Nm = DsNm
End If
DszDb = DszTny(A, Tny(A), Nm)
End Function

Function DszTny(A As Database, Tny$(), Optional DsNm$) As Ds
Dim T
For Each T In Tny
    AddDt DszTny, DtzT(A, CStr(T))
Next
End Function
Sub EnsTmpTbl(A As Database)
If HasTbl(A, "#Tmp") Then Exit Sub
A.Execute "Create Table [#Tmp] (AA Int, BB Text 10)"
End Sub

Sub RunQ(A As Database, Q)
Const CSub$ = CMod & "RunQ"
On Error GoTo X
A.Execute Q
Exit Sub
X: Dim E$: E = Err.Description: Thw CSub, "Running Sql error", "Er Sql Db", E, Q, Dbn(A)
End Sub
Sub RunQQAv(A As Database, QQ$, Av())
RunQ A, FmtQQAv(QQ, Av)
End Sub
Sub RunQQ(A As Database, QQ$, ParamArray Ap())
Dim Av(): Av = Ap
RunQQAv A, QQ, Av
End Sub

Function RszQQ(A As Database, QQ$, ParamArray Ap()) As Dao.Recordset
Dim Av(): Av = Ap
Set RszQQ = Rs(A, FmtQQAv(QQ, Av))
End Function
Function RszQ(A As Database, Q) As Dao.Recordset
Set RszQ = Rs(A, Q)
End Function
Function MovFst(A As Dao.Recordset) As Dao.Recordset
A.MoveFirst
Set MovFst = A
End Function
Function Rs(A As Database, Q) As Dao.Recordset
Const CSub$ = CMod & "Rs"
On Error GoTo X
Set Rs = A.OpenRecordset(Q)
Exit Function
X: Thw CSub, "Error in opening Rs", "Er Sql Db", Err.Description, Q, Dbn(A)
End Function

Function HasReczQ(A As Database, Q) As Boolean
HasReczQ = HasRec(Rs(A, Q))
End Function

Function HasQryz(A As Database, Q) As Boolean
HasQryz = HasReczQ(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function HasTbl(A As Database, T) As Boolean
HasTbl = HasItn(A.TableDefs, T)
End Function

Function HasTblByMSys(A As Database, T) As Boolean
HasTblByMSys = HasRec(Rs(A, FmtQQ("Select Name from MSysObjects where Type in (1,6) and Name='?'", T)))
End Function

Function IsDbOk(A As Database) As Boolean
On Error GoTo X
IsDbOk = IsStr(A.Name)
Exit Function
X:
End Function
Function DbPth$(A As Database)
DbPth = Pth(Dbn(A))
End Function

Function Qny(A As Database) As String()
Qny = SyzQ(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Function RszQry(A As Database, QryNm$) As Dao.Recordset
Set RszQry = A.QueryDefs(QryNm).OpenRecordset
End Function

Function SrcTny(A As Database) As String()
Dim T
For Each T In Tni(A)
    PushNonBlank SrcTny, A.TableDefs(T).SourceTableName
Next
End Function

Function TmpTny(A As Database) As String()
TmpTny = AywPfx(Tny(A), "#")
End Function

Function Tni(A As Database)
Asg Itr(Tny(A)), Tni
End Function

Function Tbli(A As Database)
Asg Itr(Tny(A)), Tbli
End Function

Function Tny(A As Database) As String()
Set A = Dao.DBEngine.OpenDatabase(A.Name)
Dim T As TableDef
For Each T In A.TableDefs
    If Not IsSysTd(T) Then
        If Not IsHidTd(T) Then
            PushI Tny, T.Name
        End If
    End If
Next
End Function

Function TnyzADO(A As Database) As String()
TnyzADO = TnyzFb(A.Name)
End Function


Function Tny1(A As Database) As String()
Dim T As TableDef, O$()
Dim X As Dao.TableDefAttributeEnum
X = Dao.TableDefAttributeEnum.dbHiddenObject Or Dao.TableDefAttributeEnum.dbSystemObject
For Each T In A.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI Tny1, T.Name
    End Select
Next
End Function

Function TnyzMSysObj(A As Database) As String()
TnyzMSysObj = SyzQ(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
End Function

Private Sub ZZ_Qny()
'DmpAy Qny(Db(SampFbzDutyDta))
End Sub

Private Sub Z_DszDb()
Dim A As Database, Tny0, Act As Ds, Ept As Ds
Stop
ZZ1:
    Set A = Db(SampFbzDutyDta)
    Act = DszDb(A)
    BrwDs Act
    Exit Sub
ZZ2:
    Tny0 = "Permit PermitD"
    'Set Act = Ds( Tny0)
    Stop
End Sub

Private Sub ZZ()
Dim Db As Database
Dim B As Dao.TableDef
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

Sub RenTTzAddPfx(A As Database, TT$, Pfx$)
Dim T
For Each T In Ny(TT)
    RenTblzAddPfx A, CStr(T), Pfx
Next
End Sub

Function TdStrAy(A As Database, TT$) As String()
Dim T$, I
For Each I In ItrzTT(TT)
    T = I
    PushI TdStrAy, TdStrzT(A, T)
Next
End Function

Sub CrtTblzTmp(A As Database)
DrpT A, "#Tmp"
A.TableDefs.Append TmpTd
End Sub
Sub RenTbl(A As Database, T, ToNm$)
A.TableDefs(T).Name = ToNm
End Sub

Sub RenTblzFmPfx(A As Database, FmPfx$, ToPfx$)
Dim T As TableDef
For Each T In A.TableDefs
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

Property Get TblDes$(A As Database, T)
TblDes = PrpVal(A.TableDefs(T).Properties, C_Des)
End Property

Property Let TblDes(A As Database, T, Des$)
PrpVal(A.TableDefs(T).Properties, C_Des) = Des
End Property

Property Get TblAttDes$(A As Dao.Database)
TblAttDes = TblDes(A, "Att")
End Property

Property Let TblAttDes(A As Dao.Database, Des$)
TblDes(A, "Att") = Des
End Property

Property Set TblDesDic(A As Database, D As Dictionary)
Dim T
For Each T In D.Keys
    TblDes(A, T) = D(T)
Next
End Property

Property Get FldDesDic(A As Database) As Dictionary
Dim T$, I, J, F$, D$
Set FldDesDic = New Dictionary
For Each I In Tni(A)
    T = I
    For Each J In Fny(A, T)
        F = J
        D = FldDes(A, T, F)
        If D <> "" Then FldDesDic.Add T & "." & F, D
    Next
Next
End Property

Property Set FldDesDic(A As Database, D As Dictionary)
Dim T$, F$, Des$, TDotF$, I, J
For Each I In D.Keys
    TDotF = I
    Des = D(TDotF)
    If HasDot(TDotF) Then
        AsgBrkDot TDotF, T, F
        FldDes(A, T, F) = Des
    Else
        For Each J In Tny(A)
            T = J
            If HasFld(A, T, F) Then
                FldDes(A, T, F) = Des
            End If
        Next
    End If
Next
End Property

Sub ClsDb(A As Database)
On Error Resume Next
A.Close
End Sub

Property Get TblDesDic(A As Database) As Dictionary
Dim T, O As New Dictionary
For Each T In Tni(A)
    AddDiczNonBlankStr O, T, TblDes(A, T)
Next
Set TblDesDic = O
End Property

Function JnStrDiczTwoColSql(TwoColSql) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
'Set JnStrDiczTwoColSql = JnStrDicTwoFldRs(RszSql(TwoColSql))
End Function

Private Sub ZZ_BrwTbl()
Dim D As Database
Stop
DrpTT D, "#A #B"
RunQ D, "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from [#T]"
BrwTT D, "#A #T #B"
End Sub

Function TnizInp(A As Database)
Asg Itr(TnyzInp(A)), TnizInp
End Function

Function TnyzInp(A As Database) As String()
TnyzInp = AywLik(Tny(A), ">*")
End Function

Function ReOpnDb(A As Database) As Database
Set ReOpnDb = Db(A.Name)
End Function

Function Dbn$(A As Database)
On Error GoTo X
Dbn = A.Name
Exit Function
X:
Dbn = Err.Description
End Function

Function FmtNRec(D As Database) As String()
Dim T$(): T = Tny(D)
Erase XX
X "Fb   " & Dbn(D)
X "NTbl " & Si(T)
Dim I, J%
For Each I In Itr(T)
    J = J + 1
    X AlignR(J, 3) & " " & AlignR(NReczT(D, I), 7) & " " & I
Next
FmtNRec = XX
Erase XX
End Function

Sub DmpNRec(D As Database)
Dmp FmtNRec(D)
End Sub

