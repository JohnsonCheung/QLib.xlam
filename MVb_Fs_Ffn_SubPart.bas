Attribute VB_Name = "MVb_Fs_Ffn_SubPart"
Option Explicit

Function FdrFfn$(A)
FdrFfn = Fdr(Pth(A))
End Function
Function CutPth$(Ffn)
Dim P%: P = InStrRev(Ffn, PthSep)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn, P + 1)
End Function
Function Fn$(Ffn)
Fn = CutPth(Ffn)
End Function

Function FfnUp$(Ffn)
FfnUp = ParPth(Pth(Ffn)) & Fn(Ffn)
End Function

Function Fnn$(Ffn)
Fnn = RmvExt(Fn(Ffn))
End Function

Function RmvExt$(A)
Dim B$, C$, P%
B = Fn(A)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
RmvExt = Pth(A) & C
End Function

Function Ext$(A)
Dim B$, P%
B = Fn(A)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function

Function FfnPth$(Ffn)
FfnPth = Pth(Ffn)
End Function

Function PthUp$(Pth, NUp%)
Dim O$: O = Pth
Dim J%
For J = 1 To NUp
    O = ParPth(O)
Next
PthUp = O
End Function
Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function


