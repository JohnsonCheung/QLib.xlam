Attribute VB_Name = "MXls_Fx"
Option Explicit

Function BrwFx(Fx)
WbVis WbzFx(Fx)
End Function

Sub CrtFx(A)
WbSavAs(NewWb, A).Close
End Sub

Function FxEns$(Fx)
If Not HasFfn(Fx) Then CrtFx Fx
FxEns = Fx
End Function

Function FstWsn$(Fx)
FstWsn = FstItr(WsNyzFx(Fx))
End Function

Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & CnStrzFxAdo(A)
End Function

Sub OpnFx(A)
Dim C$
C = FmtQQ("Excel ""?""", A)
Shell C, vbMaximizedFocus
End Sub

Sub FxRmvWsIfHas(A, Wsn)
If HasFxw(A, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(A)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsWbNoSav B
End If
End Sub

Function DrsFxq(A, Sql) As Drs
Set DrsFxq = DrsArs(CnzFx(A).Execute(Sql))
End Function

Sub RunFxq(Fx, Sql)
CnzFx(Fx).Execute Sql
End Sub
Function TmpDbFx(Fx$) As Database
Set TmpDbFx = TmpDbzFxww(Fx, WsNyzFx(Fx))
End Function

Function TmpDbzFxww(Fx$, WW) As Database
Dim O As Database
   Set O = TmpDb
'LnkFx O, Fx, NyzNN(WW)
Set TmpDbzFxww = O
End Function

Function WbzFx(Fx) As Workbook
Set WbzFx = Xls.Workbooks.Open(Fx)
End Function

Function WszFxw(Fx, Optional Wsn$ = "Data") As Worksheet
Set WszFxw = WszWb(WbzFx(Fx), Wsn)
End Function

Function ArsFxwf(A, W, F) As ADODB.Recordset
Set ArsFxwf = ArsCnq(CnzFx(A), SqlSel_F_Fm(F, W & "$"))
End Function

Function WsCdNyzFx(Fx) As String()
Dim Wb As Workbook
Set Wb = WbzFx(Fx)
WsCdNyzFx = WsCdNy(Wb)
Wb.Close False
End Function

Function DtzFxw(Fx, Optional Wsn0$) As Dt
Dim N$: N = FxDftWsn(Fx, Wsn0)
Dim Sql$: Sql = FmtQQ("Select * from [?$]", N)
Set DtzFxw = DtzDrs(DrsFxq(Fx, Sql), N)
End Function

Function IntAyFxwf(Fx, W, F) As Integer()
IntAyFxwf = IntAyzArs(ArsFxwf(Fx, W, F))
End Function

Function WszFxwSy(A, W, Optional F = 0) As String()
WszFxwSy = SyzArs(ArsFxwf(A, W, F))
End Function

Private Sub ZZ_WsNyzFx()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D WsNyzFx(Fx)
End Sub

Private Sub Z_FstWsn()
Dim Fx$
Fx = SampFx_KE24
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = FstWsn(Fx)
    C
    Return
End Sub

Private Sub Z_TmpDbFx()
Dim Db As Database: Set Db = TmpDbFx("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
'DmpAy TnyDb(Db)
Db.Close
End Sub

Private Sub Z_WsNyzFx()
Dim Fx$
'GoTo ZZ
GoSub T1
Exit Sub
T1:
    Fx = SampFx_KE24
    Ept = SySsl("Sheet1 Sheet21")
    GoSub Tst
    Return
Tst:
    Act = WsNyzFx(Fx)
    C
    Return
ZZ:
    DmpAy WsNyzFx(SampFx_KE24)
End Sub

Private Sub ZZ()
Dim A$
Dim B As Variant
CrtFx B
OpnFx B
'FxRmvWsIfHas B, B
DrsFxq B, B
RunFxq B, B
WbzFx B
WszFxw B, A
End Sub

Private Sub Z()
Z_FstWsn
End Sub
