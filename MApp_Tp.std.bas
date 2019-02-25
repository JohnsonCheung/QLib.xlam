Attribute VB_Name = "MApp_Tp"
Option Explicit
Const CMod$ = "MApp_Tp."
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
    Info CSub, "Tp not exist, no Import", "Tp", Tp1
    Exit Sub
End If
If IsOldAttz(AppDb(Apn), "Tp", Tp1) Then ImpAtt "Tp", Tp1 '<== Import
End Sub

'===============================================
Private Function TpFx$(Apn)
TpFx = TpPth & Apn & "(Template).xlsx"
End Function

Private Function TpFxm$(Apn)
TpFxm = TpPth & Apn & "(Template).xlsm"
End Function
Sub OpnTp(Apn)
Dim Tp1$: Tp1 = Tp(Apn): If Tp1 = "" Then Info CSub, "Tp not found", "Tp", Tp1
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

Sub ExpTpzFb(Fb$, ToFfn$)
ExpTpz Db(Fb), ToFfn
End Sub

Sub ExpTpz(Db As Database, ToFfn$)
ExpAttz Db, "Tp", ToFfn
End Sub

Sub ExpTp(ToFfn$)
ExpAtt "Tp", ToFfn
End Sub
