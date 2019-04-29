Attribute VB_Name = "MVb_Lg_Lgr"
Option Explicit

Sub LgrBrw()
BrwFt LgrFt
End Sub

Property Get LgrFilNo%()
LgrFilNo = FnoApp(LgrFt)
End Property

Property Get LgrFt$()
LgrFt = LgrPth & "Log.txt"
End Property

Sub LgrLg(Msg$)
Dim F%: F = LgrFilNo
Print #F, NowStr & " " & Msg
If LgrFilNo = 0 Then Close #F
End Sub

Property Get LgrPth$()
Dim O$:
'O = WrkPth: EnsPth O
O = O & "Log\": EnsPth O
LgrPth = O
End Property
