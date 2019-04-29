Attribute VB_Name = "MApp_Wrk"
Option Explicit
Private Db As Database
Property Get W() As Database
Set W = A
End Property
Sub WCls()
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
Function WTny() As String()
WTny = Tny(W)
End Function

Function WStru(Optional TT) As String()
WStru = StruzTT(W, TT)
End Function

Sub WDrp(TT)
DrpTT W, TT
End Sub

Sub WBrw(Apn$)
OpnFb WAcs, WFb(Apn)
WAcs.Visible = True
End Sub
Sub WKill(Apn$)
WCls
Kill WFb(Apn)
End Sub

Function WAcs() As Access.Application
Static A As New Access.Application
Set WAcs = A
End Function

Function WPth$(Apn$)
WPth = EnsPth(TmpHom & Apn)
End Function

Function WFb$(Apn$)
WFb = WPth(Apn) & Apn & "(Wrk).accdb"
End Function
Sub WRenTbl(Fm$, ToTbl$)
RenTbl W, Fm, ToTbl
End Sub

Sub WClr()
Dim T, Tny$()
Tny = WTny: If Si(Tny) = 0 Then Exit Sub
For Each T In Tny
    WDrp T
Next
End Sub

Sub WImpTbl(TT)
'ImpTbl W, TT
End Sub


Sub WDmpStru(TT$)
Dmp StruzTT(W, TT)
End Sub


