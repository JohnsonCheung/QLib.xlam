Attribute VB_Name = "QVb_Fs_Ffn_SubPart"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_SubPart."
Private Const Asm$ = "QVb"

Function FdrzFfn$(Ffn$)
FdrzFfn = Fdr(Pth(Ffn))
End Function
Function CutPth$(Ffn$)
Dim P%: P = InStrRev(Ffn$, PthSep)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn$, P + 1)
End Function
Function Fn$(Ffn$)
Fn = CutPth(Ffn$)
End Function

Function FfnUp$(Ffn$)
FfnUp = ParPth(Pth(Ffn$)) & Fn(Ffn$)
End Function

Function Fnn$(Ffn$)
Fnn = RmvExt(Fn(Ffn$))
End Function

Function RmvExt$(Ffn$)
Dim B$, C$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
RmvExt = Pth(Ffn) & C
End Function
Function IsExtInAp(Ffn$, ParamArray Ap()) As Boolean

End Function
Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = HasEle(Av, V)
End Function

Function ExtzFfn$(Ffn$)
ExtzFfn = Ext(Ffn)
End Function

Function Ext$(Ffn$)
Dim B$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function

Function FfnPth$(Ffn$)
FfnPth = Pth(Ffn$)
End Function

Function UpPth$(Pth$, NUp%)
Dim O$: O = Pth
Dim J%
For J = 1 To NUp
    O = ParPth(O)
Next
UpPth = O
End Function
Function Pth$(Ffn$)
Dim P%: P = InStrRev(Ffn$, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn$, P)
End Function


