Attribute VB_Name = "QVb_Lin_Term_FstNTerm"
Option Explicit
Private Const CMod$ = "MVb_Lin_Term_FstNTerm."
Private Const Asm$ = "QVb"

Function Fst2Term(Lin) As String()
Fst2Term = FstNTerm(Lin, 2)
End Function

Function Fst3Term(Lin) As String()
Fst3Term = FstNTerm(Lin, 3)
End Function
Function Fst4Term(Lin) As String()
Fst4Term = FstNTerm(Lin, 4)
End Function

Function FstNTerm(Lin, N%) As String()
Dim J%, L$
L = Lin
For J = 1 To N
    PushI FstNTerm, ShfT1(L)
Next
End Function
