Attribute VB_Name = "QApp_App_App"
Option Compare Text
Option Explicit
Private Const CMod$ = "App."
Private Type A
    Appn As String
    Appv As String
    AppDb As Database
End Type
Private A As A
Sub SetApp(Appn$, Appv$)
A.Appn = Appn
A.Appv = Appv
Dim Fb$: Fb = WFb
DltFfnIf Fb
CrtFb Fb
Set A.AppDb = Db(Fb)
End Sub

Function AppPth$()
Stop '
End Function
Function AppTpFx$()
AppTpFx = AppPth & A.Appn & "(Template).xlsx"
End Function

Function AppTpFxm$()
AppTpFxm = AppPth & A.Appn & "(Template).xlsm"
End Function

Sub AppOpnTp()
OpnFx AppTp
End Sub

Function AppTp$()
Dim F$
F = AppTpFxm: If HasFfn(F) Then AppTp = F: Exit Function
F = AppTpFx:  If HasFfn(F) Then AppTp = F: Exit Function
End Function

Function AppOupFxzNxt$()
AppOupFxzNxt = NxtFfnzAva(AppOupFx)
End Function

Function AppOupPth$()
AppOupPth = ValzPm(A.AppDb, "OupPth")
End Function

Function AppFfnzPm$(PmNm$)
AppFfnzPm = ValzPm(A.AppDb, PmNm & "Ffn")
End Function

Function AppOupFx$()
AppOupFx = AppOupPth & A.Appn & ".xlsx"
End Function

Function AppFb$()
'C:\Users\user\Desktop\MHD\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3.app.accdb
AppFb = AppHom & JnDotAp(A.Appn, A.Appv, "app", "accdb")
End Function

Function AppHom$()
Static Y$
If Y = "" Then Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
AppHom = Y
End Function

Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, CUsr = "User"
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function

Sub ImpAppTp()
Dim T$: T = AppTp
Const CSub$ = CMod & "ImpTp"
If T = "" Then
    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", A.AppDb, A.Appn, T, AppPth
    Exit Sub
End If
Dim D As Database: Set D = A.AppDb
If IsAttOld(D, "Tp", T) Then ImpAtt D, "Tp", T '<== Import
End Sub

Function AppTpWsCdNy() As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Function

'==============================================
Function AppTpPth$()
AppTpPth = EnsPth(TmpHom & "Template\")
End Function

Sub RfhAppTpWc()
RfhFx AppTp, AppFb
End Sub

Function AppTpWb() As Workbook
Set AppTpWb = WbzFx(AppTp)
End Function

Function WcsyzAppTp() As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(AppTp)
'TpWcsy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub ExpAppTp()
ExpAtt A.AppDb, "Tp", AppTp
End Sub

Function PthzPm$(A As Database, PmNm$)
PthzPm = EnsPthSfx(ValzPm(A, PmNm & "Pth"))
End Function

Function Pjfnm$(A As Database, PmNm$)
Pjfnm = ValzPm(A, PmNm & "Fn")
End Function

Function FfnzPm(A As Database, PmNm$)
FfnzPm = PthzPm(A, PmNm) & Pjfnm(A, PmNm)
End Function

Property Get ValzPm$(A As Database, PmNm$)
Dim Q$: Q = FmtQQ("Select ? From Pm where CUsr='?'", PmNm, CUsr)
ValzPm = ValzQ(A, Q)
End Property

Property Let ValzPm(A As Database, PmNm$, V$)
With A.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(PmNm).Value = V
    .Update
End With
End Property

Sub BrwPm(A As Database)
BrwT A, "Pm"
End Sub

Property Get W() As Database
Set W = A.AppDb
End Property
Sub WCls()
On Error Resume Next
W.Close
End Sub

Sub WRun(QQ$, ParamArray Ap())
Dim Av(): Av = Ap
RunQQAv W, QQ, Av
End Sub
Function WTny() As String()
WTny = Tny(W)
End Function

Function WStru(Optional TT$) As String()
WStru = StruzTT(W, TT)
End Function

Sub WDrp(TT$)
DrpTT W, TT
End Sub

Sub WBrw()
OpnFb WAcs, WFb
WAcs.Visible = True
End Sub
Sub WKill()
WCls
Kill WFb
End Sub

Function WAcs() As Access.Application
Static A As New Access.Application
Set WAcs = A
End Function

Function WPth$()
WPth = EnsPth(TmpHom & A.Appn)
End Function

Function WFb$()
WFb = WPth & A.Appn & "(Wrk).accdb"
End Function

Sub WRenTbl(Fm$, ToTbl$)
RenTbl W, Fm, ToTbl
End Sub

Sub WClr()
Dim T$, I, Tny$()
Tny = WTny: If Si(Tny) = 0 Then Exit Sub
For Each I In Tny
    T = I
    WDrp T
Next
End Sub

Sub WImpTbl(TT$)
'ImpTbl  W, TT
End Sub

Sub WDmpStru(TT$)
Dmp StruzTT(W, TT)
End Sub

