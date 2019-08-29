Attribute VB_Name = "QDao_B_Ado"
Option Compare Text
Option Explicit
Public Const FFoFTy$ = "T F Ty"
Public Const FFoFxwFTy$ = "Fx Ws T F Ty"
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Ado."

Function ArsCnq(Cn As AdoDb.Connection, Q) As AdoDb.Recordset
Set ArsCnq = Cn.Execute(Q)
End Function

Function ArszFbq(Fb, Q) As AdoDb.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Function ArszFxw(Fx, Wsn) As AdoDb.Recordset
Set ArszFxw = ArsCnq(CnzFx(Fx), SqlSel_T(CattnzWsn(Wsn)))
End Function

Sub ArunzFbQ(Fb, Q)
CnzFb(Fb).Execute Q
End Sub

Function Cat(A As AdoDb.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set Cat = O
End Function

Function CattnzWsn$(Wsn)
If IsNeedQte(Wsn) Then
    CattnzWsn = QteSng(Wsn & "$")
Else
    CattnzWsn = Wsn & "$"
End If
End Function

Function CatzFb(Fb) As Catalog
Set CatzFb = Cat(CnzFb(Fb))
End Function

Function CatzFx(Fx) As Catalog
Set CatzFx = Cat(CnzFx(Fx))
End Function

Function Cn(AdoCnStr) As AdoDb.Connection
Set Cn = New AdoDb.Connection
Cn.Open AdoCnStr
End Function

Function CnStrzDbt$(D As Database, T)
CnStrzDbt = DtaSrczScvl(D.TableDefs(T).Connect)
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

Function CnStrzFxAdo$(A)
'CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?", A) 'Try
CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A) 'Ok
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

Function CnzFb(A) As AdoDb.Connection
Set CnzFb = Cn(CnStrzFbzAsAdo(A))
End Function

Function CnzFx(Fx) As AdoDb.Connection
Set CnzFx = Cn(CnStrzFxAdo(Fx))
End Function

Function CvAdoTy(A) As AdoDb.DataTypeEnum
CvAdoTy = A
End Function

Function CvAxt(A) As ADOX.Table
':Axt: :Adox.Table #Catalog-Table#
Set CvAxt = A
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
    DftWny = Wny(Fx)
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

Function DftWsny(Wsny0, Fx) As String()
Dim O$()
    O = CvSy(Wsny0)
If Si(O) = 0 Then
    DftWsny = Wny(Fx)
Else
    DftWsny = O
End If
End Function

Function DoFTy(Cat As ADOX.Catalog, T) As Drs
Dim Axt As ADOX.Table, ODy()
Set Axt = Cat.Tables(T)
Dim C As Column
For Each C In Cat.Tables(T).Columns
    PushI ODy, Array(T, C.Name, ShtTyzAdo(C.Type))
Next
DoFTy = Drs(FoFTy, ODy)
End Function

Function DoFTyzFxw(Fx, W) As Drs
DoFTyzFxw = DoFTy(CatzFx(Fx), CattnzWsn(W))
End Function

Function DoFxwFTy(Fx, W) As Drs
Dim A As Drs
A = DoFTyzFxw(Fx, W)
DoFxwFTy = InsColzDrsCC(A, "Fx W", Fx, W)
End Function

Function DrsCnq(Cn As AdoDb.Connection, Q) As Drs
DrsCnq = DrszArs(ArsCnq(Cn, Q))
End Function

Function DrsFbqAdo(Fb, Q) As Drs
DrsFbqAdo = DrszArs(ArszFbq(Fb, Q))
End Function

Function DrszArs(A As AdoDb.Recordset) As Drs
DrszArs = Drs(FnyzArs(A), DyoArs(A))
End Function

Function DrszFxw(Fx, Wsn) As Drs
DrszFxw = DrszArs(ArszFxw(Fx, Wsn))
End Function

Function DrzAfds(A As AdoDb.Fields, Optional N%) As Variant()
Dim F As AdoDb.Field
For Each F In A
   PushI DrzAfds, F.Value
Next
End Function

Function DtaSrczScvl(Scvl$)
DtaSrczScvl = VzScvl(Scvl, "Data Source")
End Function

Function DWsfzFxw(Fx, W) As Drs
Stop
'DWsfzFxw = DrszFF("Fx Wsn F Ty", ODy)
End Function

Function DyoArs(A As AdoDb.Recordset) As Variant()
While Not A.EOF
    PushI DyoArs, DrzAfds(A.Fields)
    A.MoveNext
Wend
End Function

Function FFzFxw$(Fx, Wsn)
FFzFxw = TLin(FnyzFxw(Fx, Wsn))
End Function

Function FnyzAfds(A As AdoDb.Fields) As String()
FnyzAfds = Itn(A)
End Function

Function FnyzArs(A As AdoDb.Recordset) As String()
FnyzArs = FnyzAfds(A.Fields)
End Function

Function FnyzCatt(Cat As ADOX.Catalog, T) As String()
':Catt: :Cat,T
Dim Axt As ADOX.Table
Set Axt = Cat.Tables(T)
FnyzCatt = Itn(Cat.Tables(T).Columns)
End Function

Function FnyzFbt(Fb, T) As String()
FnyzFbt = Fny(Db(Fb), T)
End Function

Function FnyzFbtAdo(Fb, T) As String()
Dim C As ADOX.Catalog
Set C = CatzFb(Fb)
FnyzFbtAdo = FnyzCatt(C, T)
End Function

Function FnyzFfnTn(Ffn, T$) As String()
Const CSub$ = CMod & "FnyzFfnTn"
Select Case True
Case IsFx(Ffn): FnyzFfnTn = FnyzFxw(Ffn, T)
Case IsFb(Ffn): FnyzFfnTn = FnyzFbt(Ffn, T)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn T", Ffn, T
End Select
End Function

Function FnyzFxw(Fx, W) As String()
FnyzFxw = FnyzCatt(CatzFx(Fx), CattnzWsn(W))
End Function

Function FoFTy() As String()
FoFTy = SyzSS(FFoFTy)
End Function

Function HasFbt(Fb, T) As Boolean
HasFbt = HasEle(TnyzFb(Fb), T)
End Function

Function HasFxw(Fx, Wsn) As Boolean
HasFxw = HasEle(Wny(Fx), Wsn)
End Function

Function HasReczArs(A As AdoDb.Recordset) As Boolean
HasReczArs = Not NoReczArs(A)
End Function

Private Function HasTblzCT(A As Catalog, T) As Boolean
HasTblzCT = HasItn(A.Tables, T)
End Function

Function HasTblzFfnT(Ffn, T) As Boolean
Const CSub$ = CMod & "HasTblzFfnTn"
Select Case True
Case IsFx(Ffn): HasTblzFfnT = HasFxw(Ffn, T)
Case IsFx(Ffn): HasTblzFfnT = HasFxw(Ffn, T)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn T", Ffn, T
End Select
End Function

Function IntAyzArs(A As AdoDb.Recordset, Optional Col = 0) As Integer()
IntAyzArs = IntoColzArs(EmpIntAy, A, Col)
End Function

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

Function NoReczArs(A As AdoDb.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoReczArs = True
End Function

Sub RunCnSqy(Cn As AdoDb.Connection, Sqy$())
Dim Q
For Each Q In Itr(Sqy)
   Cn.Execute Q
Next
End Sub

Function SyzArs(A As AdoDb.Recordset, Optional Col = 0) As String()
SyzArs = IntoColzArs(EmpSy, A, Col)
End Function

Function TnyzCat(A As Catalog) As String()
TnyzCat = Itn(A.Tables)
End Function

Function TnyzFb(Fb) As String()
TnyzFb = Tny(Db(Fb))
End Function

Function TnyzFbByAdo(Fb) As String()
'TnyzAdoFb = TnyzCat(CatzFb(Fb))
TnyzFbByAdo = AeKss(TnyzCat(CatzFb(Fb)), "MSys* f_*_Data")
End Function

Function VzScvl(Scvl$, Nm$)
VzScvl = Bet(EnsSfx(Scvl, ";"), Nm & "=", ";")
End Function

Function Wny(Fx, Optional InclAllOtherTbl As Boolean) As String()
Dim Tny$(), T
Tny = TnyzCat(CatzFx(Fx))
If InclAllOtherTbl Then
    Wny = Tny
    Exit Function
End If
For Each T In Itr(Tny)
    PushNB Wny, WsnzCattn(T)
Next
End Function

Function WnyzWb(A As Workbook) As String()
WnyzWb = Wny(A.FullName)
End Function

Function WsnzCattn$(Cattn)
If HasSfx(Cattn, "FilterDatabase") Then Exit Function
WsnzCattn = RmvSfx(RmvSngQte(Cattn), "$")
End Function

Private Sub Z()
Z_TnyzFb

MAdoX_Cat:
End Sub

Private Sub Z_ArunzFbQ()
Const Fb$ = SampFbzDutyPgm
Const Q$ = "Select * into [#a] from Permit"
DrpFbt Fb, "#a"
ArunzFbQ Fb, Q
End Sub

Private Sub Z_Cn()
Dim O As AdoDb.Connection
Set O = Cn(GetCnStr_ADO_SampSQL_EXPR_NOT_WRK)
Stop
End Sub

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

Private Sub Z_CnzFb()
Dim Cn
Set Cn = CnzFb(SampFbzDutyDta)
Stop
End Sub

Private Sub Z_DrsCnq()
Dim Cn As AdoDb.Connection: Set Cn = CnzFx(SampFxzKE24)
Dim Q$: Q = "Select * from [Sheet1$]"
WszDrs DrsCnq(Cn, Q)
End Sub

Private Sub Z_DrsFbqAdo()
Const Fb$ = SampFbzDutyDta
Const Q$ = "Select * from Permit"
BrwDrs DrsFbqAdo(Fb, Q)
End Sub

Private Sub Z_DyArs()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
BrwDy DyoArs(ArsCnq(CnzFb(SampFbzDutyDta), Q))
End Sub

Private Sub Z_TnyzFb()
DmpAy TnyzFbByAdo(SampFbzDutyDta)
End Sub

Private Sub Z_Wny()
D Wny(SampFxzKE24)
End Sub

Private Sub Z_Wny2()
Const Fx$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
D Wny(Fx)
End Sub
