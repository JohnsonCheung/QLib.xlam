Attribute VB_Name = "QApp_App_SalRpt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MApp_SalRpt."
Private Const Asm$ = "QApp"
Public Const DoczSrp = "Sales-Report-Parameter."
Property Get DftSrpDic() As Dictionary
Dim X As Boolean, Y As New Dictionary
If Not X Then
    X = True
    With Y
        .Add "BrkCrd", False
        .Add "BrkDiv", False
        .Add "BrkMbr", False
        .Add "BrkSto", False
        .Add "LisCrd", ""
        .Add "LisSto", ""
        .Add "LisDiv", ""
        .Add "FmDte", "20170101"
        .Add "ToDte", "20170131"
        .Add "SumLvl", "M"
        .Add "InclAdr", False
        .Add "InclNm", False
        .Add "InclPhone", False
        .Add "InclEmail", False
    End With
End If
Set DftSrpDic = Y
End Property

