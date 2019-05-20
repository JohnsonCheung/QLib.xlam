Attribute VB_Name = "QVb_Str_Apd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Apd."
Private Const Asm$ = "QVb"
Private Const NS$ = "MVb_Str"
Function ApdCrLf$(S)
ApdCrLf = ApdIf(S, vbCrLf)
End Function
Function PpdSpcIf$(S)
PpdSpcIf = PpdIf(S, " ")
End Function
Function ApdIf$(S, Sfx$)
If S = "" Then ApdIf = S: Exit Function
ApdIf = S & Sfx
End Function
Function ApdIfzAy(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    PushI ApdIfzAy, ApdIf(I, Sfx)
Next
End Function
Function PpdIf$(S, Pfx$)
If S = "" Then PpdIf = S: Exit Function
PpdIf = Pfx & S
End Function

