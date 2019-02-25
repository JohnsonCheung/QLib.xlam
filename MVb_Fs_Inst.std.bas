Attribute VB_Name = "MVb_Fs_Inst"
Option Explicit
Const CMod$ = "MVb_Fs_Ffn."
Function FfnInst$(Ffn)
FfnInst = PthInst(Pth(Ffn)) & Fn(Ffn)
End Function

Function PthInst$(Pth)
PthInst = AddFdrEns(Pth, NowStr)
End Function

Function CrtPthzInst$(Pth)
CrtPthzInst = PthInst(Pth)
End Function
