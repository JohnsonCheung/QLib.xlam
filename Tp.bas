VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CMod$ = ""
Private Type A
    App As App
End Type
Private A As A
Friend Function Init(App As App) As Tp
Set A.App = App
Set Init = Me
End Function
Function TpFn$() 'Fst Fn in Tbl.Fld.Ssk-Att.Att.Tp
End Function
Function HasTp() As Boolean
HasTp = HasFfn(Tp)
End Function

Sub Import()
Const CSub$ = CMod & "ImpTp"
If Tp = "" Then
'    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", App.Db .Nm, Tp, App.WPth
    Exit Sub
End If
Dim D As Database: Set D = App.Db
If IsOldAtt(D, "Tp", Tp) Then ImpAtt D, "Tp", Tp '<== Import
End Sub
Property Get Tp$()

End Property
Property Get TpWsCdNy() As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Property

'==============================================
Property Get TpPth$()
TpPth = EnsPth(TmpHom & "Template\")
End Property

Sub RfhWcTp()
RfhFx Tp, App.Fb
End Sub

Function TpWb() As Workbook
Set TpWb = WbzFx(Tp)
End Function

Function TpWcSy(Apn) As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(Tp)
TpWcSy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub Export()
ExpAtt App.Db, "Tp", Tp
End Sub



