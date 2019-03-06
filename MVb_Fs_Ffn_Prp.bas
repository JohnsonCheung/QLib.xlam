Attribute VB_Name = "MVb_Fs_Ffn_Prp"
Option Explicit
Function FfnDte(Ffn) As Date
FfnDte = FileDateTime(Ffn)
End Function
Function FfnSz&(Ffn)
If Not HasFfn(Ffn) Then FfnSz = -1: Exit Function
FfnSz = FileLen(Ffn)
End Function
Function FfnFdr$(Ffn)
FfnFdr = Fdr(Pth(Ffn))
End Function
Function TimFfn(Ffn) As Date
If HasFfn(Ffn) Then TimFfn = FfnDte(Ffn)
End Function

Function SzDotDTimFfn$(A)
If HasFfn(A) Then SzDotDTimFfn = DteTimStr(FfnDte(A)) & "." & FfnSz(A)
End Function

Sub AsgTimFfnSz(A$, OTim As Date, OSz&)
If Not HasFfn(A) Then
    OTim = 0
    OSz = 0
    Exit Sub
End If
OTim = TimFfn(A)
OSz = FfnSz(A)
End Sub

Function FfnTimStr$(A)
If HasFfn(A) Then
    FfnTimStr = DteTimStr(FfnDte(A))
End If
End Function


