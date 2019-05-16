Attribute VB_Name = "QXls_Fx"
Option Explicit
Private Const CMod$ = "MXls_Fx."
Private Const Asm$ = "QXls"

Function BrwFx(Fx)
If HasFfn(Fx) Then Debug.Print "No Fx:" & Fx
ShwWb WbzFx(Fx)
End Function

Sub CrtFx(Fx)
WbSavAs(NewWb, Fx).Close
End Sub

Function FxEns$(Fx)
If Not HasFfn(Fx) Then CrtFx Fx
FxEns = Fx
End Function

Function FstWsn$(Fx)
FstWsn = FstItm(Wny(Fx))
End Function

Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & CnStrzFxAdo(A)
End Function
Function HasFx(Fx) As Boolean
Dim Wb As Workbook
For Each Wb In Xls.Workbooks
    If Wb.FullName = Fx Then HasFx = True: Exit Function
Next
End Function

Function OpnFx(Fx) As Workbook
ThwIf_FfnNotExist Fx, CSub
Set OpnFx = Xls.Workbooks.Open(Fx)
End Function

Sub RmvWsIf(Fx, Wsn$)
If HasFxw(Fx, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(Fx)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsWbNoSav B
End If
End Sub

Function DrszFxq(Fx, Q) As Drs
DrszFxq = DrszArs(CnzFx(Fx).Execute(Q))
End Function

Sub RunFxqByCn(Fx, Q)
CnzFx(Fx).Execute Q
End Sub
Function TmpDbzFx(Fx) As Database
Set TmpDbzFx = TmpDbzFxWny(Fx, Wny(Fx))
End Function

Function TmpDbzFxWny(Fx, Wny$()) As Database
Dim O As Database
   Set O = TmpDb
Dim W
For Each W In Itr(Wny)
    LnkFx O, ">" & W, Fx, W
Next
Set TmpDbzFxWny = O
End Function

Function Wb(Fx) As Workbook
Set Wb = Xls.Workbooks(Fx)
End Function
Function WbzFx(Fx) As Workbook
Set WbzFx = Wb(Fx)
End Function

Function WszFxw(Fx, Optional Wsn$ = "Data") As Worksheet
Set WszFxw = WszWb(WbzFx(Fx), Wsn)
End Function

Function ArszFxwf(Fx, W$, F$) As AdoDb.Recordset
Set ArszFxwf = ArsCnq(CnzFx(Fx), SqlSel_F_T(F, W & "$"))
End Function

Function WsCdNyzFx(Fx) As String()
Dim Wb As Workbook
Set Wb = WbzFx(Fx)
WsCdNyzFx = WsCdNy(Wb)
Wb.Close False
End Function

Function DtzFxw(Fx, Optional Wsn0$) As Dt
Dim N$: N = DftWsn(Wsn0, Fx)
Dim Q$: Q = FmtQQ("Select * from [?$]", N)
DtzFxw = DtzDrs(DrszFxq(Fx, Q), N)
End Function

Function IntAyFxwf(Fx, W$, F$) As Integer()
IntAyFxwf = IntAyzArs(ArszFxwf(Fx, W, F))
End Function

Function SyzFxwf(Fx, W$, F$) As String()
SyzFxwf = SyzArs(ArszFxwf(Fx, W, F))
End Function

Private Sub ZZ_Wny()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D Wny(Fx)
End Sub

Private Sub Z_FstWsn()
Dim Fx$
Fx = SampFxzKE24
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = FstWsn(Fx)
    C
    Return
End Sub

Private Sub Z_TmpDbzFx()
Dim Db As Database: Set Db = TmpDbzFx("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
DmpAy Tny(Db)
Db.Close
End Sub

Private Sub Z_Wny()
Dim Fx$
'GoTo ZZ
GoSub T1
Exit Sub
T1:
    Fx = SampFxzKE24
    Ept = SyzSS("Sheet1 Sheet21")
    GoSub Tst
    Return
Tst:
    Act = Wny(Fx)
    C
    Return
ZZ:
    DmpAy Wny(SampFxzKE24)
End Sub

Private Sub ZZ()
Z_FstWsn
End Sub
