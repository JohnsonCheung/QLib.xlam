Attribute VB_Name = "QVb_Str_Align"
Option Explicit
Private Const CMod$ = "MVb_Str_Align."
Private Const Asm$ = "QVb"
Function Align$(V, W%)
Dim S$: S = V
If IsStr(V) Then
    Align = AlignL(S, W)
Else
    Align = AlignR(S, W)
End If
End Function

Function AlignL$(S$, W%)
Dim L%: L = Len(S)
If L >= W Then
    AlignL = S
Else
    AlignL = S & Space(W - Len(S))
End If
End Function

Function AlignR$(S$, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

