Attribute VB_Name = "MVb_Str_Align"
Option Explicit

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

