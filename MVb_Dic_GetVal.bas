Attribute VB_Name = "MVb_Dic_GetVal"
Option Explicit

Function VyzDicKK(Dic As Dictionary, Ky$()) As Variant()
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then
        PushI VyzDicKK, Dic(K)
    End If
Next
End Function

