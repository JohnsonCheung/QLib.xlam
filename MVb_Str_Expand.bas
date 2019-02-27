Attribute VB_Name = "MVb_Str_Expand"
Option Explicit
Function Expand$(QVbl$, ExpandByTLin$)
Dim T, O$(), L$
L = RplVbl(QVbl)
For Each T In TermAy(ExpandByTLin)
    PushI O, RplQ(L, T)
Next
Expand = JnCrLf(O)
End Function
Private Sub Z_Expand()
Dim QVbl$
Erase XX
X "Sub Push?(O() As ?, M As ?)"
X "Dim N&"
X "N = ?Sz(O)"
X "ReDim Preserve O(N)"
X "O(N) = M"
X "End Sub"
X ""
X "Function ?Sz&(A() As ?)"
X "On Error Resume Next"
X "?Sz = Ubound(A) + 1"
X "End Function"
X ""
QVbl = JnVBar(XX)
Brw Expand(QVbl, "S1S2 XX")
End Sub

