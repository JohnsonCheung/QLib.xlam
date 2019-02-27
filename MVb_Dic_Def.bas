Attribute VB_Name = "MVb_Dic_Def"
Option Explicit
Function DefDic(Ly$(), KK) As Dictionary
Dim L, S As Aset, T1$, Rst$, O As New Dictionary
Set S = TermAset(KK)
If S.Has("*Er") Then Thw CSub, "KK cannot have Term-*Er", "KK Ly", KK, Ly
For Each L In Ly
    AsgTRst L, T1, Rst
    If S.Has(T1) Then
        PushItmzSyDic O, T1, Rst
    Else
'        PushItmzSyDic , O, L
    End If
    Set DefDic = O
Next
End Function
