Attribute VB_Name = "MDao_Ado"
Option Explicit
Function CvTc(A) As ADOX.Table
Set CvTc = A
End Function

Function TnyzCat(A As Catalog) As String()
TnyzCat = Itn(A.Tables)
End Function

Function CatzFb(Fb) As Catalog
Set CatzFb = CatCn(CnzFb(Fb))
End Function

Function CatCn(A As ADODB.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CatCn = O
End Function

Function CatzFx(Fx) As Catalog
Set CatzFx = CatCn(CnzFx(Fx))
End Function

Function FnyCatTbl(Cat As ADOX.Catalog, T) As String()
Dim CT As ADOX.Table
Set CT = Cat.Tables(T)
FnyCatTbl = Itn(Cat.Tables(T).Columns)
End Function

Function DrsFxw(Fx, Wsn) As DRs
Set DrsFxw = DrsArs(ArsFxw(Fx, Wsn))
End Function
Function ArsFxw(Fx, Wsn) As ADODB.Recordset
Set ArsFxw = ArsCnq(CnzFx(Fx), SqlSel_Fm(CatT(Wsn)))
End Function
Sub RunCnSqy(A As ADODB.Connection, Sqy$())
Dim Q
For Each Q In Itr(Sqy)
   A.Execute Q
Next
End Sub

Private Sub Z_DrsCnq()
Dim Cn As ADODB.Connection: Set Cn = CnzFx(SampFx_KE24)
Dim Q$: Q = "Select * from [Sheet1$]"
WszDrs DrsCnq(Cn, Q)
End Sub

Function ArsCnq(A As ADODB.Connection, Q) As ADODB.Recordset
Set ArsCnq = A.Execute(Q)
End Function

Function DrsCnq(A As ADODB.Connection, Q) As DRs
Set DrsCnq = DrsArs(ArsCnq(A, Q))
End Function
Function DrsFbqAdo(A$, Q$) As DRs
Set DrsFbqAdo = DrsArs(ARsFbq(A, Q))
End Function

Private Sub Z_DrsFbqAdo()
Const Fb$ = SampFbzDutyDta
Const Q$ = "Select * from Permit"
BrwDrs DrsFbqAdo(Fb, Q)
End Sub

Function ARsFbq(Fb$, Q$) As ADODB.Recordset
Set ARsFbq = CnzFb(Fb).Execute(Q)
End Function

Function DrsArs(A As ADODB.Recordset) As DRs
Set DrsArs = DRs(FnyArs(A), DryArs(A))
End Function

Function DryArs(A As ADODB.Recordset) As Variant()
While Not A.EOF
    PushI DryArs, DrzAfds(A.Fields)
    A.MoveNext
Wend
End Function

Private Sub Z_DryArs()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
BrwDry DryArs(ArsCnq(CnzFb(SampFbzDutyDta), Q))
End Sub

Function FnyArs(A As ADODB.Recordset) As String()
FnyArs = FnyAfds(A.Fields)
End Function

Function IntAyzARs(A As ADODB.Recordset, Optional Col = 0) As Integer()
IntAyzARs = IntoColzArs(A, EmpIntAy, Col)
End Function

Private Function HasCatT(A As Catalog, T) As Boolean
HasCatT = HasItn(A.Tables, T)
End Function

Private Sub Z_TnyzFb()
DmpAy TnyzFb(SampFbzDutyDta)
End Sub

Private Sub Z_WsNyzFx()
DmpAy WsNyzFx(SampFx_KE24)
End Sub
Function HasTblzFfnTblNm(Ffn, TblNm) As Boolean
Select Case True
Case IsFx(Ffn): HasTblzFfnTblNm = HasFxw(Ffn, TblNm)
Case IsFx(Ffn): HasTblzFfnTblNm = HasFxw(Ffn, TblNm)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn TblNm", Ffn, TblNm
End Select
End Function
Function HasFbt(Fb, T) As Boolean
HasFbt = HasEle(TnyzFb(Fb), T)
End Function

Function HasFxw(Fx, W) As Boolean
HasFxw = HasEle(WsNyzFx(Fx), W)
End Function

Function TnyzFb(Fb) As String()
TnyzFb = TnyzAdoFb(Fb)
End Function

Function TnyzAdoFb(Fb) As String()
TnyzAdoFb = AyeLikss(TnyzCat(CatzFb(Fb)), "MSys* f_*_Data")
'TnyzAdoFb = TnyzCat(CatzFb(Fb))
End Function

Function WsNyzFx(Fx) As String()
Dim T
For Each T In Itr(TnyzCat(CatzFx(Fx)))
    PushNonBlankStr WsNyzFx, WsnzCatT(T)
Next
End Function

Function FnyzFbt(Fb, T) As String()
Dim C As ADOX.Catalog
Set C = CatzFb(Fb)
FnyzFbt = FnyCatTbl(C, T)
End Function


Private Sub Z_CnStrzFbAdo()
Dim CnStr$
'
CnStr = CnStrzFbAdo(SampFbzDutyDta)
GoSub Tst
'
CnStr = CnStrzFbAdo(CurrentDb.Name)
'GoSub Tst
Exit Sub
Tst:
    Cn(CnStr).Close
    Return
End Sub
Private Sub Z_Cn()
Dim O As ADODB.Connection
Set O = Cn(GetCnStr_ADO_SampSQL_EXPR_NOT_WRK)
Stop
End Sub

Function Cn(AdoCnStr) As ADODB.Connection
Set Cn = New ADODB.Connection
Cn.Open AdoCnStr
End Function

Function DftWsNy(WsNy0, Fx$) As String()
Dim O$()
    O = CvSy(WsNy0)
If Sz(O) = 0 Then
    DftWsNy = WsNyzFx(Fx)
Else
    DftWsNy = O
End If
End Function
Function DftTny(Tny0, Fb$) As String()
Dim O$()
    O = CvSy(Tny0)
If Sz(O) = 0 Then
    DftTny = TnyzFb(Fb)
Else
    DftTny = O
End If
End Function
Function FxDftWsNy(A, WsNy0) As String()
Dim O$(): O = CvSy(WsNy0)
If Sz(O) = 0 Then
    FxDftWsNy = WsNyzFx(A)
Else
    FxDftWsNy = O
End If
End Function

Function FxDftWsn$(A, Wsn0$)
If Wsn0 = "" Then
    FxDftWsn = FstWsn(A)
    Exit Function
End If
FxDftWsn = Wsn0
End Function

Function CnStrzFbAdo$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'CnStrzFbAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;User ID=Admin;Mode=Share Deny None;"
'Locking Mode=1 means page (or record level) according to https://www.spreadsheet1.com/how-to-refresh-pivottables-without-locking-the-source-workbook.html
'The ADO connection object initialization property which controls how the database is locked, while records are being read or modified is: Jet OLEDB:Database Locking Mode
'Please note:
'The first user to open the database determines the locking mode to be used while the database remains open.
'A database can only be opened is a single mode at a time.
'For Page-level locking, set property to 0
'For Row-level locking, set property to 1
'With 'Jet OLEDB:Database Locking Mode = 0', the source spreadshseet is locked, while PivotTables update. If the property is set to 1, the source file is not locked. Only individual records (Table rows) are locked sequentially, while data is being read.
CnStrzFbAdo = FmtQQ(C, A)
End Function

Function CnStrzFxAdo$(A)
'CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?", A) 'Try
CnStrzFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A) 'Ok
End Function

Function DtaSrczScl(DtaSrcScl$)
DtaSrczScl = TakBet(DtaSrcScl, "Data Source=", ";")
End Function
Function DtaSrc$(A As Database, T)
DtaSrc = DtaSrczScl(A.TableDefs(T).Connect)
End Function

Function CnzFx(Fx) As ADODB.Connection
Set CnzFx = Cn(CnStrzFxAdo(Fx))
End Function

Function CnzFb(A) As ADODB.Connection
Set CnzFb = Cn(CnStrzFbAdo(A))
End Function

Private Sub Z_CnzFb()
Dim Cn
Set Cn = CnzFb(SampFbzDutyDta)
Stop
End Sub
Function FFzFxw$(Fx$, Wsn$)
FFzFxw = TLin(FnyzFxw(Fx, Wsn))
End Function
Function FnyzFfnTblNm(Ffn, TblNm) As String()
Select Case True
Case IsFx(Ffn): FnyzFfnTblNm = FnyzFxw(Ffn, TblNm)
Case IsFb(Ffn): FnyzFfnTblNm = FnyzFbt(Ffn, TblNm)
Case Else: Thw CSub, "Ffn must be Fx or Fb", "Ffn TblNm", Ffn, TblNm
End Select
End Function
Function FnyzFxw(Fx, W) As String()
Dim Cat As ADOX.Catalog
Set Cat = CatzFx(Fx)
FnyzFxw = FnyCatTbl(Cat, CatT(W))
End Function
Function CvAdoTy(A) As ADODB.DataTypeEnum
CvAdoTy = A
End Function

Function CatT$(Wsn)
If IsNeedQuote(Wsn) Then
    CatT = QuoteSng(Wsn & "$")
Else
    CatT = Wsn & "$"
End If
End Function

Function WsnzCatT$(CatT)
Dim I$
I = RmvSngQuote(CatT)
If LasChr(I) = "$" Then WsnzCatT = RmvLasChr(I)
End Function

Private Sub Z()
Z_TnyzFb

MAdoX_Cat:
End Sub

Function IntoColzArs(A As ADODB.Recordset, Into, Optional Col = 0)
IntoColzArs = AyCln(Into)
With A
    While Not .EOF
        PushI IntoColzArs, Nz(.Fields(Col).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
End Function

Function SyzArs(A As ADODB.Recordset, Optional Col = 0) As String()
SyzArs = IntoColzArs(A, EmpSy, Col)
End Function

Sub ArunzFbq(A$, Q$)
CnzFb(A).Execute Q
End Sub

Private Sub Z_ArunzFbq()
Const Fb$ = SampFbzDuty_Pgm
Const Q$ = "Select * into [#a] from Permit"
DrpvFbt Fb, "#a"
ArunzFbq Fb, Q
End Sub

Function DrzAfds(A As ADODB.Fields, Optional N%) As Variant()
Dim F As ADODB.Field
For Each F In A
   PushI DrzAfds, F.Value
Next
End Function

Function FnyAfds(A As ADODB.Fields) As String()
FnyAfds = Itn(A)
End Function


