Attribute VB_Name = "QApp_App_App"
Option Compare Text
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
TpFx = WPth(A) & A.Nm & "(Template).xlsx"
End Function

Function TpFxm$(A As App)
TpFxm = WPth(A) & A.Nm & "(Template).xlsm"
End Function
Function WFb$(A As App)

End Function
Sub OpnTp(A As App)
OpnFx Tp(A)
End Sub

Function Tp$(A As App)
Dim F$
F = TpFxm(A): If HasFfn(F) Then Tp = F: Exit Function
F = TpFx(A):  If HasFfn(F) Then Tp = F: Exit Function
End Function

Function OupFxzNxt$(A As App)
OupFxzNxt = NxtFfnzAva(OupFx(A))
End Function

Function OupPth$(A As App)
OupPth = ValzPm(A.Db, "OupPth")
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

Function HomzApp$(A As App)
Static Y$
If Y = "" Then Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
HomzApp = Y
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

Sub ImportTp(A As App)
Dim T$: T = Tp(A)
Const CSub$ = CMod & "ImpTp"
If T = "" Then
    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", A.Db, A.Nm, T, WPth(A)
    Exit Sub
End If
Dim D As Database: Set D = A.Db
If IsOldAtt(D, "Tp", T) Then ImpAtt D, "Tp", T '<== Import
End Sub

Function TpWsCdNy(A As App) As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Function

'==============================================
Function TpPth$()
TpPth = EnsPth(TmpHom & "Template\")
End Function

Sub RfhTpWc(A As App)
RfhFx Tp(A), WFb(A)
End Sub

Function TpWb(A As App) As Workbook
Set TpWb = WbzFx(Tp(A))
End Function

Function TpWcsy(A As App) As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(Tp(A))
'TpWcsy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub ExpTp(A As App)
ExpAtt A.Db, "Tp", Tp(A)
End Sub


