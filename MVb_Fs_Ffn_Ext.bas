Attribute VB_Name = "MVb_Fs_Ffn_Ext"
Option Explicit
Function RplExt$(Ffn$, NewExt$)
RplExt = RmvExt(Ffn$) & NewExt
End Function


