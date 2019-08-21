Attribute VB_Name = "QVb_F_VbLg"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Lg_Lgr."
Private Const Asm$ = "QVb"

Sub LgrBrw()
BrwFt LgrFt
End Sub

Property Get LgrFilNo%()
LgrFilNo = FnoA(LgrFt)
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
