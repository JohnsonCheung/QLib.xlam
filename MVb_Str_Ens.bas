Attribute VB_Name = "MVb_Str_Ens"
Option Explicit
Function SfxEns$(S$, Sfx$)
If HasSfx(S, Sfx) Then SfxEns = S: Exit Function
SfxEns = S & Sfx
End Function

Function SfxDotEns$(S$)
SfxDotEns = SfxEns(S, ".")
End Function

Function SfxSemiEns$(S$)
SfxSemiEns = SfxEns(S, ";")
End Function
