Attribute VB_Name = "MVb_Lin_Term_NTermRst"
Option Explicit
Function SyzTRst(Lin) As String()
SyzTRst = SyzNTermRst(Lin, 1)
End Function

Function Syz2TRst(Lin) As String()
Syz2TRst = SyzNTermRst(Lin, 2)
End Function

Function Syz3TRst(Lin) As String()
Syz3TRst = SyzNTermRst(Lin, 3)
End Function

Function Syz4TRst(Lin) As String()
Syz4TRst = SyzNTermRst(Lin, 4)
End Function

Function SyzNTermRst(Lin, N%) As String()
Dim L$, J%
L = Lin
For J = 1 To N
    PushI SyzNTermRst, ShfT(L)
Next
PushI SyzNTermRst, L
End Function

Private Sub Z_SyzNTermRst()
Dim Lin$
Lin = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
Lin = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
Lin = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = T1(Lin)
    C
    Return
End Sub

Private Sub Z()
End Sub
