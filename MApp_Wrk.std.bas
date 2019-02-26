Attribute VB_Name = "MApp_Wrk"
Option Explicit
Private A As Database, Apn1$
Property Get W() As Database
Set W = A
End Property
Sub ClsWDb()
Apn1 = ""
On Error Resume Next
A.Close
End Sub
Sub OpnWDb(Apn$)
If Apn1 <> Apn Then
    Apn1 = Apn
    Set A = Db(WFb(Apn))
End If
End Sub
Sub WRun(QQ, ParamArray Ap())
Dim Av(): Av = Ap
RunQQ A, QQ, Av
End Sub
Sub WDrp(TT)
DrpTT W, TT
End Sub
Sub WBrw(Apn$)
OpnFbz WAcs, WFb(Apn)
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

