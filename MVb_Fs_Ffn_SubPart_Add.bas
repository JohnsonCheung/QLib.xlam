Attribute VB_Name = "MVb_Fs_Ffn_SubPart_Add"
Option Explicit

Function FfnAddTimSfx$(Ffn)
FfnAddTimSfx = FfnAddFnSfx(Ffn, Format(Now, "(HHMMSS)"))
End Function
Function FfnAddFnPfx$(A$, Pfx$)
FfnAddFnPfx = Pth(A) & Pfx & Fn(A)
End Function

Function FfnAddFnSfx$(Ffn, Sfx$)
FfnAddFnSfx = RmvExt(Ffn) & Sfx & Ext(Ffn)
End Function

