Attribute VB_Name = "QApp_App_App"
Option Explicit
Private Const CMod$ = "App."
Type App
    Nm As String
    Ver As String
    Db As Database
End Type

Function WPth$(A As App)
Stop '
End Function
Function TpFx$(A As App)
TpFx = WPth & A.Nm & "(Template).xlsx"
End Function

Function TpFxm$(A As App)
TpFxm = WPth & A.Nm & "(Template).xlsm"
End Function
Function WFb$(A As App)

End Function
Sub OpnTp(A As App)
OpnFx Tp
End Sub

Function Tp$(A As App)
Dim A$
A = TpFxm: If HasFfn(A) Then Tp = A: Exit Function
A = TpFx:  If HasFfn(A) Then Tp = A: Exit Function
End Function

Function OupFxzNxt$(A As App)
OupFxzNxt = NxtFfnzAva(PmOupFx)
End Function
Function OupPth$(A As App)
OupPth = ValzPm(Db, "OupPth")
End Function
Function FfnzPm$(A As App, PmNm$)
FfnzPm = ValzPm(A.Db, PmNm & "Ffn")
End Function

Function OupFx$(A As App)
OupFx = OupPth(A) & A.Nm & ".xlsx"
End Function

Function FbzApp$(A As App)
'C:\Users\user\Desktop\MHD\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3.app.accdb
FbzApp = HomzApp(A) & JnDotAp(A.Nm, A.Ver, "app", "accdb")
End Function

Function Hom$()
Static Y$
If Y = "" Then Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
Hom = Y
End Function

Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function

Sub ImportTp(A As App)
Const CSub$ = CMod & "ImpTp"
If Tp = "" Then
'    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", App.Db .Nm, Tp, App.WPth
    Exit Sub
End If
Dim D As Database: Set D = App.Db
If IsOldAtt(D, "Tp", Tp) Then ImpAtt D, "Tp", Tp '<== Import
End Sub

Function TpWsCdNy(A As App) As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Function

'==============================================
Function TpPth$()
TpPth = EnsPth(TmpHom & "Template\")
End Function

Sub RfhTpWc(A As App)
RfhFx Tp, App.Fb
End Sub

Function TpWb(A As App) As Workbook
Set TpWb = WbzFx(Tp)
End Function

Function TpWcsy(A As App) As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(Tp)
'TpWcsy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub ExpTp(A As App)
ExpAtt App.Db, "Tp", Tp
End Sub




