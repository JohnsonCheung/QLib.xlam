Attribute VB_Name = "QVb_Fs_Inst"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Inst."

Function FfnInst$(Ffn)
FfnInst = InstPth(Pth(Ffn)) & Fn(Ffn)
End Function

Function InstPth$(Pth)
InstPth = AddFdrEns(Pth, NowStr)
End Function

Function InstFdr$(Fdr)
InstFdr = AddFdrEns(TmpFdr(Fdr), NowStr)
End Function
Function CrtPthzInst$(Pth)
CrtPthzInst = InstPth(Pth)
End Function

Function IsInstFfn(Ffn) As Boolean
IsInstFfn = IsInstFdr(FdrzFfn(Ffn))
End Function

Function IsInstFdr(Fdr$) As Boolean
IsInstFdr = IsDteTimStr(Fdr)
End Function
