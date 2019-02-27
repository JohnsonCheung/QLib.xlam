Attribute VB_Name = "MVb_Align"
Option Explicit

Function AlignL$(A, W)
Dim L%: L = Len(A)
If L >= W Then
    AlignL = A
Else
    AlignL = A & Space(W - Len(A))
End If
End Function

Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

