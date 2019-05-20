Attribute VB_Name = "QXls_Fx"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Fx."
Private Const Asm$ = "QXls"


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

Function ChkFxww(Fx, Wsnn$, Optional FxKd$ = "Excel file") As String()
Dim W$, I
'If Not HasFfn(Fx) Then ChkFxww = MsgzMisFfn(Fx, FxKd): Exit Function
For Each I In Ny(Wsnn)
    W = I
    PushIAy ChkFxww, ChkWs(Fx, W, FxKd)
Next
End Function

Function ChkWs(Fx, Wsn, FxKd$) As String()
If HasFxw(Fx, Wsn) Then Exit Function
Dim M$
M = FmtQQ("? does not have expected worksheet", FxKd)
ChkWs = LyzFunMsgNap(CSub, M, "Folder File Expected-Worksheet Worksheets-in-file", Pth(Fx), Fn(Fx), Wsn, Wny(Fx))
End Function

Function ChkFxw(Fx, Wsn, Optional FxKd$ = "Excel file") As String()
ChkFxw = ChkHasFfn(Fx, FxKd): If Si(ChkFxw) > 0 Then Exit Function
ChkFxw = ChkWs(Fx, Wsn, FxKd)
End Function
Function ChkLnkWs(A As Database, T, Fx, Wsn, Optional FxKd$ = "Excel file") As String()
Const CSub$ = CMod & "ChkLnkWs"
Dim O$()
    O = ChkFxw(Fx, Wsn, FxKd)
    If Si(O) > 0 Then
        ChkLnkWs = O
        Exit Function
    End If
On Error GoTo X
LnkFx A, T, Fx, Wsn
Exit Function
X: ChkLnkWs = _
    LyzMsgNap("Error in linking Xls file", "Er LnkFx LnkWs ToDb AsTbl", Err.Description, Fx, Wsn, Dbn(A), T)
End Function


