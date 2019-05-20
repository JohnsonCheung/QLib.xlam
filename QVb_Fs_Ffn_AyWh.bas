Attribute VB_Name = "QVb_Fs_Ffn_AyWh"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_AyWh."
Private Const Asm$ = "QVb"
Function FxAyFfny(Ffny$()) As String()
Dim Ffn
For Each Ffn In Itr(Ffny)
    If IsFx(Ffn) Then PushI FxAyFfny, Ffn
Next
End Function

Function FbAyFfny(Ffny$()) As String()
Dim Ffn
For Each Ffn In Itr(Ffny)
    If IsFb(Ffn) Then PushI FbAyFfny, Ffn
Next
End Function

