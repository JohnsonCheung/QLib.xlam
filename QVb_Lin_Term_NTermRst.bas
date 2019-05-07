Attribute VB_Name = "QVb_Lin_Term_NTermRst"
Option Explicit
Private Const CMod$ = "MVb_Lin_Term_NTermRst."
Private Const Asm$ = "QVb"
Function SyzTRst(Lin) As String()
SyzTRst = SyzNTermRst(Lin, 1)
End Function

Function SyzN2tRst(Lin) As String()
SyzN2tRst = SyzNTermRst(Lin, 2)
End Function

Function SyzN3TRst(Lin) As String()
SyzN3TRst = SyzNTermRst(Lin, 3)
End Function

Function SyzN4tRst(Lin) As String()
SyzN4tRst = SyzNTermRst(Lin, 4)
End Function

Function SyzNTermRst(Lin, N%) As String()
Dim L$, J%
L = Lin
For J = 1 To N
    PushI SyzNTermRst, ShfT1(L)
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
