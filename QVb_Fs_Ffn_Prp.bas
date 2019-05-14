Attribute VB_Name = "QVb_Fs_Ffn_Prp"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Prp."
Private Const Asm$ = "QVb"
Function DtezFfn(Ffn) As Date
If HasFfn(Ffn) Then DtezFfn = FileDateTime(Ffn)
End Function
Function SizFfn&(Ffn)
If Not HasFfn(Ffn) Then SizFfn = -1: Exit Function
SizFfn = FileLen(Ffn)
End Function
Function SiDotDTim$(Ffn)
If HasFfn(Ffn) Then SiDotDTim = DteTimStr(DtezFfn(Ffn)) & "." & SizFfn(Ffn)
End Function

Sub AsgTimSi(Ffn, OTim As Date, OSz&)
OTim = DtezFfn(Ffn)
OSz = SizFfn(Ffn)
End Sub

Function DteTimStrzFfn$(Ffn)
DteTimStrzFfn = DteTimStr(DtezFfn(Ffn))
End Function


