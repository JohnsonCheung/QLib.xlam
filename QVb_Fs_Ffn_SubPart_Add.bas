Attribute VB_Name = "QVb_Fs_Ffn_SubPart_Add"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_SubPart_Add."
Private Const Asm$ = "QVb"

Function FfnAddTimSfx$(Ffn$)
FfnAddTimSfx = AddFnSfx(Ffn$, Format(Now, "(HHMMSS)"))
End Function
Function FfnAddFnPfx$(A$, Pfx$)
FfnAddFnPfx = Pth(A) & Pfx & Fn(A)
End Function

Function AddFnSfx$(Ffn$, Sfx$)
AddFnSfx = RmvExt(Ffn$) & Sfx & Ext(Ffn$)
End Function

