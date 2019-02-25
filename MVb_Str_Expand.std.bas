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
QVbl = "Function ?() As Lof?(): ? = A_?: End Function"
Brw Expand(QVbl, LofKK)
End Sub

