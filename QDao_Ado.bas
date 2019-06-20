Attribute VB_Name = "QDao_Ado"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Ado."
Function CvTc(A) As ADOX.Table
Set CvTc = A
End Function
Function NoReczAdo(A As AdoDb.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoReczAdo = True
End Function
Function HasReczAdo(A As AdoDb.Recordset) As Boolean
HasReczAdo = Not NoReczAdo(A)
End Function

Function TnyzCat(A As Catalog) As String()
TnyzCat = Itn(A.Tables)
End Function

Function CatzFb(Fb) As Catalog
Set CatzFb = CatCn(CnzFb(Fb))
End Function

Function CatCn(A As AdoDb.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CatCn = O
End Function

Function CatzFx(Fx) As Catalog
Set CatzFx = CatCn(CnzFx(Fx))
End Function

Function FnyzCatT(Cat As ADOX.Catalog, T) As String()
Dim CT As ADOX.Table
Set CT = Cat.Tables(T)
FnyzCatT = Itn(Cat.Tables(T).Columns)
End Function

Function DFxwFTy(Fx, W) As Drs
Dim A As Drs
A = DFTyzFxw(Fx, W)
DFxwFTy = InsColzDrsCC(A, "Fx W", Fx, W)
End Function
Function DFTyzFxw(Fx, W) As Drs
DFTyzFxw = DFTy(CatzFx(Fx), CatTnzWsn(W))
End Function

Function DFTy(Cat As ADOX.Catalog, T) As Drs
Dim CT As ADOX.Table, ODry()
Set CT = Cat.Tables(T)
Dim C As Column
For Each C In Cat.Tables(T).Columns
    PushI ODry, Array(C.Name, ShtTyzAdo(C.Type))
Next
DFTy = DrszFF("T Ty", ODry)
End Function

Function DrszFxw(Fx, Wsn) As Drs
DrszFxw = DrszArs(ArszFxw(Fx, Wsn))
End Function
Function ArszFxw(Fx, Wsn) As AdoDb.Recordset
Set ArszFxw = ArsCnq(CnzFx(Fx), SqlSel_T(CatTnzWsn(Wsn)))
End Function
Sub RunCnSqy(Cn As AdoDb.Connection, Sqy$())
Dim Q
For Each Q In Itr(Sqy)
   Cn.Execute Q
Next
End Sub

Private Sub Z_DrsCnq()
Dim Cn As AdoDb.Connection: Set Cn = CnzFx(SampFxzKE24)
Dim Q$: Q = "Select * from [Sheet1$]"
WszDrs DrsCnq(Cn, Q)
End Sub

Function ArsCnq(Cn As AdoDb.Connection, Q) As AdoDb.Recordset
Set ArsCnq = Cn.Execute(Q)
End Function

Function DrsCnq(Cn As AdoDb.Connection, Q) As Drs
DrsCnq = DrszArs(ArsCnq(Cn, Q))
End Function
Function DrsFbqAdo(Fb, Q) As Drs
DrsFbqAdo = DrszArs(ArszFbq(Fb, Q))
End Function

Private Sub Z_DrsFbqAdo()
Const Fb$ = SampFbzDutyDta
Const Q$ = "Select * from Permit"
BrwDrs DrsFbqAdo(Fb, Q)
End Sub

Function ArszFbq(Fb, Q) As AdoDb.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Function DrszArs(A As AdoDb.Recordset) As Drs
DrszArs = Drs(FnyzArs(A), DryzArs(A))
End Function

Function DryzArs(A As AdoDb.Recordset) As Variant()
While Not A.EOF
    PushI DryzArs, DrzAfds(A.Fields)
    A.MoveNext
Wend
End Function

Private Sub Z_DryArs()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
BrwDry DryzArs(ArsCnq(CnzFb(SampFbzDutyDta), Q))
End Sub

Function FnyzArs(A As AdoDb.Recordset) As String()
FnyzArs = FnyzAfds(A.Fields)
End Function

Function IntAyzArs(A As AdoDb.Recordset, Optional Col = 0) As Integer()
IntAyzArs = IntoColzArs(EmpIntAy, A, Col)
End Function

Private Function HasTblzCT(A As Catalog, T) As Boolean
HasTblzCT = HasItn(A.Tables, T)
End Function

Private Sub Z_TnyzFb()
DmpAy TnyzFbByAdo(SampFbzDutyDta)
End Sub

Private Sub Z_Wny()
D WNy(SampFxzKE24)
End Sub

Function HasTblzFfnT(Ffn, T) As Boolean
Const CSub$ = CMod & "HasTblzFfnTblNm"
Select Case True
Case IsFx(Ffn): HasTblzFfnT = HasFxw(Ffn, T)
Case IsFx(Ffn): HasTblzFfnT = HasFxw(Ffn, T)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn T", Ffn, T
End Select
End Function

Function HasFbt(Fb, T) As Boolean
HasFbt = HasEle(TnyzFb(Fb), T)
End Function

Function HasFxw(Fx, Wsn) As Boolean
HasFxw = HasEle(WNy(Fx), W)
End Function

Function TnyzFb(Fb) As String()
TnyzFb = Tny(Db(Fb))
End Function

Function TnyzFbByAdo(Fb) As String()
'TnyzAdoFb = TnyzCat(CatzFb(Fb))
TnyzFbByAdo = AyeLikss(TnyzCat(CatzFb(Fb)), "MSys* f_*_Data")
End Function

Function WnyzWb(A As Workbook) As String()
WnyzWb = WNy(A.FullName)
End Function
Private Sub Z_Wny2()
Const Fx$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
D WNy(Fx)
End Sub
Function WNy(Fx, Optional InclAllOtherTbl As Boolean) As String()
Dim Tny$(), T
Tny = TnyzCat(CatzFx(Fx))
If InclAllOtherTbl Then
    WNy = Tny
    Exit Function
End If
For Each T In Itr(Tny)
    PushNB WNy, WsnzCatTn(T)
Next
End Function

Function FnyzFbt(Fb, T) As String()
FnyzFbt = Fny(Db(Fb), T)
End Function

Function FnyzFbtAdo(Fb, T) As String()
Dim C As ADOX.Catalog
Set C = CatzFb(Fb)
FnyzFbtAdo = FnyzCatT(C, T)
End Function


Private Sub Z_CnStrzFbzAsAdo()
Dim CnStr$
'
CnStr = CnStrzFbzAsAdo(SampFbzDutyDta)
GoSub Tst
'
CnStr = CnStrzFbzAsAdo(CurrentDb.Name)
'GoSub Tst
Exit Sub
Tst:
    Cn(CnStr).Close
    Return
End Sub
Private Sub Z_Cn()
Dim O As AdoDb.Connection
Set O = Cn(GetCnStr_ADO_SampSQL_EXPR_NOT_WRK)
Stop
End Sub

Function Cn(AdoCnStr) As AdoDb.Connection
Set Cn = New AdoDb.Connection
Cn.Open AdoCnStr
End Function

Function DftWsny(Wsny0, Fx) As String()
Dim O$()
    O = CvSy(Wsny0)
If Si(O) = 0 Then
    DftWsny = WNy(Fx)
Else
    DftWsny = O
End If
End Function
Function DftTny(Tny0, Fb) As String()
If IsMissing(Tny0) Then
    DftTny = TnyzFb(Fb)
Else
    DftTny = CvSy(Tny0)
End If
End Function

Function DftWny(Wny0, Fx) As String()
If IsMissing(Wny0) Then
    DftWny = WNy(Fx)
Else
    DftWny = CvSy(Wny0)
End If
End Function

Function DftWsn$(Wsn0$, Fx)
If Wsn0 = "" Then
    DftWsn = FstWsn(Fx)
    Exit Function
End If
DftWsn = Wsn0
End Function

Function CnStrzFbzAsAdo$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'CnStrzFbzAsAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;User ID=Admin;Mode=Share Deny None;"
'Locking Mode=1 means page (or record level) according to https://www.spreadsheet1.com/how-to-refresh-pivottables-without-locking-the-source-workbook.html
'The ADO connection object initialization property which controls how the database is locked, while records are being read or modified is: Jet OLEDB:Database Locking Mode
'Please note:
'The first user to open the database determines the locking mode to be used while the database remains open.
'A database can only be opened is a single mode at a time.
'For Page-level locking, set property to 0
'For Row-level locking, set property to 1
'With 'Jet OLEDB:Database Locking Mode = 0', the source spreadshseet is locked, while PivotTables update. If the property is set to 1, the source file is not locked. Only individual records (Table rows) are locked sequentially, while data is being read.
CnStrzFbzAsAdo = FmtQQ(C, A)
End Function
Function CnStrzFxzOrFb$(Fx_or_Fb)
Dim A$: A = Fx_or_Fb
Select Case True
Case IsFx(A): CnStrzFxzOrFb = CnStrzFbzAsAdo(A)
Case IsFb(A): CnStrzFxzOrFb = CnStrzFxAdo(A)
Case IsFb(A):
Case Else: Thw CSub, "Must be either Fx or Fb", "Fx_or_Fb", A
End Select
End Function
Function CnStrzFxAdo$(A)
'CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?", A) 'Try
CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A) 'Ok
End Function
Function ValzScvl(Scvl$, Nm$)
ValzScvl = Bet(EnsSfx(Scvl, ";"), Nm & "=", ";")
End Function

Function DtaSrczScvl(Scvl$)
DtaSrczScvl = ValzScvl(Scvl, "Data Source")
End Function
Function CnStrzDbt$(A As Database, T)
CnStrzDbt = DtaSrczScvl(A.TableDefs(T).Connect)
End Function

Function CnzFx(Fx) As AdoDb.Connection
Set CnzFx = Cn(CnStrzFxAdo(Fx))
End Function

Function CnzFb(A) As AdoDb.Connection
Set CnzFb = Cn(CnStrzFbzAsAdo(A))
End Function

Private Sub Z_CnzFb()
Dim Cn
Set Cn = CnzFb(SampFbzDutyDta)
Stop
End Sub
Function FFzFxw$(Fx, Wsn)
FFzFxw = TLin(FnyzFxw(Fx, Wsn))
End Function
Function FnyzFfnTblNm(Ffn, TblNm$) As String()
Const CSub$ = CMod & "FnyzFfnTblNm"
Select Case True
Case IsFx(Ffn): FnyzFfnTblNm = FnyzFxw(Ffn, TblNm$)
Case IsFb(Ffn): FnyzFfnTblNm = FnyzFbt(Ffn, TblNm$)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn TblNm", Ffn, TblNm
End Select
End Function
Function DWsfzFxw(Fx, W) As Drs
Stop
'DWsfzFxw = DrszFF("Fx Wsn F Ty", ODry)
End Function
Function FnyzFxw(Fx, W) As String()
FnyzFxw = FnyzCatT(CatzFx(Fx), CatTnzWsn(W))
End Function
Function CvAdoTy(A) As AdoDb.DataTypeEnum
CvAdoTy = A
End Function

Function CatTnzWsn$(Wsn)
If IsNeedQte(Wsn) Then
    CatTnzWsn = QteSng(Wsn & "$")
Else
    CatTnzWsn = Wsn & "$"
End If
End Function

Function WsnzCatTn$(CatTn)
If HasSfx(CatTn, "FilterDatabase") Then Exit Function
WsnzCatTn = RmvSfx(RmvSngQte(CatTn), "$")
End Function

Private Sub Z()
Z_TnyzFb

MAdoX_Cat:
End Sub

Function IntoColzArs(IntoCol, A As AdoDb.Recordset, Optional Col = 0)
IntoColzArs = ResiU(IntoCol)
With A
    While Not .EOF
        PushI IntoColzArs, Nz(.Fields(Col).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
End Function

Function SyzArs(A As AdoDb.Recordset, Optional Col = 0) As String()
SyzArs = IntoColzArs(EmpSy, A, Col)
End Function

Sub ArunzFbQ(Fb, Q)
CnzFb(Fb).Execute Q
End Sub

Private Sub Z_ArunzFbQ()
Const Fb$ = SampFbzDutyPgm
Const Q$ = "Select * into [#a] from Permit"
DrpFbt Fb, "#a"
ArunzFbQ Fb, Q
End Sub

Function DrzAfds(A As AdoDb.Fields, Optional N%) As Variant()
Dim F As AdoDb.Field
For Each F In A
   PushI DrzAfds, F.Value
Next
End Function

Function FnyzAfds(A As AdoDb.Fields) As String()
FnyzAfds = Itn(A)
End Function


