Attribute VB_Name = "QVb_Fs_Ffn_Ext"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Ext."
Private Const Asm$ = "QVb"
Function RplExt$(Ffn, NewExt)
RplExt = RmvExt(Ffn) & NewExt
End Function


