Attribute VB_Name = "MApp_Tp"
Option Explicit
Const CMod$ = "MApp_Tp."
Function TpFn$(Apn) 'Fst Fn in Tbl.Fld.Ssk-Att.Att.Tp
End Function
Function Tp$(Apn)
Dim A$
A = TpFxm(Apn): If HasFfn(A) Then Tp = A: Exit Function
A = TpFx(Apn):  If HasFfn(A) Then Tp = A: Exit Function
End Function
Function HasTp(Apn) As Boolean
If HasFfn(TpFxm(Apn)) Then HasTp = True
HasTp = HasFfn(TpFx(Apn))
End Function

Sub ImpTp(Apn)
Const CSub$ = CMod & "ImpTp"
Dim Tp1$: Tp1 = Tp(Apn)
If Tp1 = "" Then
    Inf CSub, "Tp not exist AppFb, no Import", "AppFb Tp", AppFb(Apn), Tp1
    Exit Sub
End If
Dim D As Database: Set D = AppDb(Apn)
If IsOldAtt(D, "Tp", Tp1) Then ImpAtt D, "Tp", Tp1 '<== Import
End Sub

'===============================================
Private Function TpFx$(Apn)
TpFx = TpPth & Apn & "(Template).xlsx"
End Function

Private Function TpFxm$(Apn)
TpFxm = TpPth & Apn & "(Template).xlsm"
End Function
Sub OpnTp(Apn)
Dim Tp1$: Tp1 = Tp(Apn): If Tp1 = "" Then Inf CSub, "Tp not found", "Tp", Tp1
OpnFx Tp1
End Sub

Property Get TpWsCdNy() As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Property

'==============================================
Property Get TpPth$()
TpPth = PthEns(TmpHom & "Template\")
End Property

Sub RfhTp(Apn)
WbVis RfhWb(TpWb(Apn), AppFb(Apn))
End Sub

Sub RfhWcTp(Apn)
RfhFx Tp(Apn), AppFb(Apn)
End Sub

Function TpWb(Apn) As Workbook
Set TpWb = WbzFx(Tp(Apn))
End Function

Function TpWcSy(Apn) As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(Tp(Apn))
TpWcSy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub ExpTp(Apn$, ToFfn$)
ExpAtt AppDb(Apn), "Tp", ToFfn
End Sub

