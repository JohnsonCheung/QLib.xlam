Attribute VB_Name = "MVb_Str_Apd"
Option Explicit
Private Const Ns$ = "MVb_Str"
Function ApdCrLf$(S$)
ApdCrLf = ApdIf(S, vbCrLf)
End Function
Function PpdSpcIf$(S$)
PpdSpcIf = PpdIf(S, " ")
End Function
Function ApdIf$(S$, Sfx$)
If S = "" Then ApdIf = S: Exit Function
ApdIf = S & Sfx
End Function
Function PpdIf$(S$, Pfx$)
If S = "" Then PpdIf = S: Exit Function
PpdIf = Pfx & S
End Function

