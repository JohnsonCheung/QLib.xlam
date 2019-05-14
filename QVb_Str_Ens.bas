Attribute VB_Name = "QVb_Str_Ens"
Option Explicit
Private Const CMod$ = "MVb_Str_Ens."
Private Const Asm$ = "QVb"

Function EnsSfxDot$(S)
EnsSfxDot = EnsSfx(S, ".")
End Function

Function EnsSfxSemi$(S)
EnsSfxSemi = EnsSfx(S, ";")
End Function
