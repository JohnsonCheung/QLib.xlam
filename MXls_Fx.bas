Attribute VB_Name = "MXls_Fx"
Option Explicit

Function BrwFx(Fx$)
WbVis WbzFx(Fx$)
End Function

Sub CrtFx(Fx$)
WbSavAs(NewWb, Fx).Close
End Sub

Function FxEns$(Fx$)
If Not HasFfn(Fx$) Then CrtFx Fx
FxEns = Fx
End Function

Function FstWsn$(Fx$)
FstWsn = FstItm(Wny(Fx$))
End Function

Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & CnStrzFxAdo(A)
End Function

Sub OpnFx(Fx$)
ThwIfFfnNotExist Fx, CSub
Dim C$
C = FmtQQ("Excel ""?""", Fx)
Shell C, vbMaximizedFocus
End Sub

Sub RmvWsIf(Fx$, Wsn$)
If HasFxw(Fx, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(Fx)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsWbNoSav B
End If
End Sub

Function DrszFxq(Fx$, Q$) As Drs
DrszFxq = DrszArs(CnzFx(Fx).Execute(Q))
End Function

Sub RunFxq(Fx$, Sql)
CnzFx(Fx$).Execute Sql
End Sub
Function TmpDbFx(Fx$) As Database
Set TmpDbFx = TmpDbzFxww(Fx$, Wny(Fx$))
End Function

Function TmpDbzFxww(Fx$, WW) As Database
Dim O As Database
   Set O = TmpDb
'LnkFx O, Fx, TermAy(WW)
Set TmpDbzFxww = O
End Function


Function Wb(Fx$) As Workbook
Set Wb = Xls.Workbooks(Fx)
End Function
Function WbzFx(Fx$) As Workbook
Set WbzFx = Wb(Fx)
End Function

Function WszFxw(Fx$, Optional Wsn$ = "Data") As Worksheet
Set WszFxw = WszWb(WbzFx(Fx$), Wsn)
End Function

Function ArszFxwf(Fx$, W$, F$) As AdoDb.Recordset
Set ArszFxwf = ArsCnq(CnzFx(Fx), SqlSel_F_T(F, W & "$"))
End Function

Function WsCdNyzFx(Fx$) As String()
Dim Wb As Workbook
Set Wb = WbzFx(Fx$)
WsCdNyzFx = WsCdNy(Wb)
Wb.Close False
End Function

Function DtzFxw(Fx$, Optional Wsn0$) As Dt
Dim N$: N = DftWsn(Fx$, Wsn0)
Dim Sql$: Sql = FmtQQ("Select * from [?$]", N)
DtzFxw = DtzDrs(DrszFxq(Fx$, Sql), N)
End Function

Function IntAyFxwf(Fx$, W$, F$) As Integer()
IntAyFxwf = IntAyzArs(ArszFxwf(Fx$, W, F))
End Function

Function SyzFxwf(Fx$, W$, F$) As String()
SyzFxwf = SyzArs(ArszFxwf(Fx, W, F))
End Function

Private Sub ZZ_Wny()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D Wny(Fx$)
End Sub

Private Sub Z_FstWsn()
Dim Fx$
Fx = SampFxzKE24
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = FstWsn(Fx$)
    C
    Return
End Sub

Private Sub Z_TmpDbFx()
Dim Db As Database: Set Db = TmpDbFx("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
'DmpAy TnyDb(Db)
Db.Close
End Sub

Private Sub Z_Wny()
Dim Fx$
'GoTo ZZ
GoSub T1
Exit Sub
T1:
    Fx = SampFxzKE24
    Ept = SySsl("Sheet1 Sheet21")
    GoSub Tst
    Return
Tst:
    Act = Wny(Fx$)
    C
    Return
ZZ:
    DmpAy Wny(SampFxzKE24)
End Sub

Private Sub Z()
Z_FstWsn
End Sub
