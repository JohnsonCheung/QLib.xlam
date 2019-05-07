Attribute VB_Name = "QVb_Fs_Inst"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Inst."

Function FfnInst$(Ffn$)
FfnInst = PthInst(Pth(Ffn$)) & Fn(Ffn$)
End Function

Function PthInst$(Pth)
PthInst = AddFdrEns(Pth, NowStr)
End Function

Function CrtPthzInst$(Pth)
CrtPthzInst = PthInst(Pth)
End Function

Function IsInstFfn(Ffn$) As Boolean
IsInstFfn = IsInstFdr(FfnFdr(Ffn$))
End Function

Function IsInstFdr(Fdr$) As Boolean
IsInstFdr = IsDteTimStr(Fdr)
End Function
