Attribute VB_Name = "QVb_Dic_GetVal"
Option Explicit
Private Const CMod$ = "MVb_Dic_GetVal."
Private Const Asm$ = "QVb"

Function VyzDicKK(Dic As Dictionary, Ky$()) As Variant()
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then
        PushI VyzDicKK, Dic(K)
    End If
Next
End Function

