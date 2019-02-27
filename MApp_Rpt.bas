Attribute VB_Name = "MApp_Rpt"
Option Explicit
Sub GenOupFx(Apn$, OupFx$)
Dim Fb$: Fb = AppFb(Apn)
ExpTpzFb Fb, OupFx
RfhWb(WbzFx(OupFx), Fb).Save
With Xls
    .Visible = True
    .WindowState = xlMaximized
    Interaction.AppActivate .Caption
End With
End Sub

