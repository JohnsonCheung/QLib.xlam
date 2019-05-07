Attribute VB_Name = "QVb_Fs_Ffn_AyWh"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_AyWh."
Private Const Asm$ = "QVb"
Function FxAyFfnAy(FfnSy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnSy)
    If IsFx(Ffn$) Then PushI FxAyFfnAy, Ffn
Next
End Function

Function FbAyFfnAy(FfnSy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnSy)
    If IsFb(Ffn$) Then PushI FbAyFfnAy, Ffn
Next
End Function

