Attribute VB_Name = "MApp_Wrk"
Option Explicit
Private A As Database, Apn1$

Property Get W() As Database
Set W = A
End Property

Sub WCls()
Apn1 = ""
On Error Resume Next
A.Close
End Sub
Sub WIniOpn(Apn$)
WIni Apn
WOpn Apn
End Sub
Sub WIni(Apn$)
Dim Fb$: Fb = WFb(Apn)
DltFfnIf Fb
CrtFb Fb
End Sub
Sub WOpn(Apn$)
Set A = Db(WFb(Apn))
End Sub

Sub WRun(QQ, ParamArray Ap())
Dim Av(): Av = Ap
RunQQAv A, QQ, Av
End Sub

Sub WDrp(TT)
DrpTT W, TT
End Sub

Sub WBrw(Apn$)
OpnFb WAcs, WFb(Apn)
WAcs.Visible = True
End Sub

Function WAcs() As Access.Application
Static A As New Access.Application
Set WAcs = A
End Function

Function WPth$(Apn)
WPth = PthEns(TmpHom & Apn)
End Function

Function WFb$(Apn)
WFb = WPth(Apn) & Apn & "(Wrk).accdb"
End Function
