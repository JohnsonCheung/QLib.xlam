Attribute VB_Name = "MApp_Wrk"
Option Explicit
Function WDb(Apn$) As Database
Static X As Boolean, Y As Database
Dim Fb$: Fb = WFb(Apn)
If Not X Then
    X = True
    DltFfnIf Fb
    EnsFb Fb
    Set Y = Db(Fb)
End If
Set WDb = Y
End Function

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

