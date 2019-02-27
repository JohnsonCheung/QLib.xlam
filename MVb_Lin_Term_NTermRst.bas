Attribute VB_Name = "MVb_Lin_Term_NTermRst"
Option Explicit
Function SyTRst(Lin) As String()
SyTRst = SyNTermRst(Lin, 1)
End Function

Function Sy2TRst(Lin) As String()
Sy2TRst = SyNTermRst(Lin, 2)
End Function

Function Sy3TRst(Lin) As String()
Sy3TRst = SyNTermRst(Lin, 3)
End Function

Function Sy4TRst(Lin) As String()
Sy4TRst = SyNTermRst(Lin, 4)
End Function

Function SyNTermRst(Lin, N%) As String()
Dim L$, J%
L = Lin
For J = 1 To N
    PushI SyNTermRst, ShfT(L)
Next
PushI SyNTermRst, L
End Function

Private Sub Z_SyNTermRst()
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
