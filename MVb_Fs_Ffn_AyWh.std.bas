Attribute VB_Name = "MVb_Fs_Ffn_AyWh"
Option Explicit
Function FxAyFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If IsFx(Ffn) Then PushI FxAyFfnAy, Ffn
Next
End Function

Function FbAyFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If IsFb(Ffn) Then PushI FbAyFfnAy, Ffn
Next
End Function

